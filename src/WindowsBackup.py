from shutil import copy2, rmtree
from subprocess import call, DEVNULL, STDOUT
import os, urllib.request, sys, zipfile, ctypes
from datetime import datetime


if getattr(sys, 'frozen', False):
	script_dir = os.path.dirname(sys.executable)
else:
	script_dir = os.path.dirname(os.path.abspath(__file__))

abs_input_path = os.path.join(script_dir, "profile.txt")
abs_logs_path = os.path.join(script_dir, "logs.txt")

signature = os.environ.get('HOMEPATH')[os.environ.get('HOMEPATH').rindex("\\")+1:] + "_" + str(datetime.now()).replace(" ", "_").replace(":", "-")
# locks = []


def Main():
	global abs_input_path
	global abs_logs_path
	global signature
	# global locks

	dst = ""
	srcs = []

	print("\n\n<--- NEW SESSION: \"" + signature + "\"  --->\n", file=open(abs_logs_path, "a"))

	if not isAdmin():
		print("\n--> Running without admin rights. Exiting.\n", file=open(abs_logs_path, "a"))
		if len(sys.argv) == 1:
			print("\n--> Running without admin rights. Exiting.\n")
		return

	if len(sys.argv) == 1:
		print("\n\n")

		while dst.strip() == "":
			dst = input("\nDestination Folder -> ")
		print("\n")

		src = ""
		while src.strip() == "":
			src = input("Source Folder -> ")
			if src.lower() == "outlook":
				src = "C:" + os.environ.get('HOMEPATH').replace("\\", "/") + "/Documents/Outlook Files"
		srcs.append(src)

		while True:
			src = input("Source Folder -> ")
			if src:
				if src.lower() == "outlook":
					src = "C:" + os.environ.get('HOMEPATH').replace("\\", "/") + "/Documents/Outlook Files"
				srcs.append(src)
			else:
				break
	else:
		try:
			f = open(abs_input_path, "r")
		except FileNotFoundError:
			print("\n--> File \"profile.txt\" could not be found. Exiting\n", file=open(abs_logs_path, "a"))
			print("\n<----- Done ----->\n", file=open(abs_logs_path, "a"))
			sys.exit()
		i = 0
		for line in f:
			if line.strip() != "":
				if i == 0:
					dst = line.rstrip('\n')
					i += 1
				else:
					if line.rstrip('\n').lower() == "outlook":
						line = "C:" + os.environ.get('HOMEPATH').replace("\\", "/") + "/Documents/Outlook Files"
					srcs.append(line.rstrip('\n'))
		f.close()

	print("Destination: \"" + dst + "\"\n", file=open(abs_logs_path, "a"))

	for src in srcs:
		print("Source: \"" + src + "\"\n", file=open(abs_logs_path, "a"))

	for src in srcs:
		mainCopy(src, dst)

	dst = dst.replace('\\', '/')

	# while(len(locks) > 0):
		# print(len(locks))

	print("\nCompressing files to: \"" + str(dst + dst[dst.rindex('/'):] + '_' + signature + '.zip') + "\"", file=open(abs_logs_path, "a"))

	if len(sys.argv) == 1:
		print("\nCompressing files to: \"" + str(dst + dst[dst.rindex('/'):] + '_' + signature + '.zip') + "\"")

	try:
		zipf = zipfile.ZipFile(dst + dst[dst.rindex('/'):] + '_' + signature + '.zip', 'w', zipfile.ZIP_DEFLATED)
		zipdir(dst, zipf)
		zipf.close()
	except FileNotFoundError:
		print("\n--> ABORTING: No such file or directory: \"" + dst + dst[dst.rindex('/'):] + '_' + signature + '.zip\"', file=open(abs_logs_path, "a"))
		if len(sys.argv) == 1:
			print("\n--> ABORTING: No such file or directory: \"" + dst + dst[dst.rindex('/'):] + '_' + signature + '.zip\"')
		return

	for item in os.listdir(dst):
		if os.path.isdir(os.path.join(dst, item)):
			rmtree(os.path.join(dst, item))
		elif not str(item).endswith('.zip'):
			os.remove(os.path.join(dst, item))

	print("\n<----- Done ----->\n", file=open(abs_logs_path, "a"))

	if len(sys.argv) == 1:
		print("\nDone")


def isAdmin():
	try:
		is_admin = os.geteuid == 0
	except AttributeError:
		is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
	return is_admin


def zipdir(path, ziph):
	global signature

	for root, dirs, files in os.walk(path):
		for file in files:
			if not str(file).endswith('.zip'):
				ziph.write(os.path.join(root, file))


def mainCopy(src, dst):
	global abs_logs_path
	# global locks

	src = src.replace('\\', '/').replace("//", "/").replace("\"", "")
	dst = dst.replace('\\', '/').replace("//", "/").replace("\"", "")
	
	if src.endswith('/'):
		src = src[:len(src) - 1]

	if dst.endswith('/'):
		dst = dst[:len(dst) - 1]

	if not os.path.exists(src):
		print("\n--> Source doesn't exist: \"" + str(src) + "\"", file=open(abs_logs_path, "a"))
		if len(sys.argv) == 1:
			print("\n--> Source doesn't exist: \"" + str(src) + "\"")
		return

	if os.path.isdir(src):
		dst += src[src.rindex('/'):].replace('\\', '/')
		copytree(src, dst)
	else:
		try:
			if not src.endswith(".pst"):
				copy2(src, dist)
				print("\nCopied: \"" + str(src) + "\"", file=open(abs_logs_path, "a"))
			else:
				try:
					file_name, headers = urllib.request.urlretrieve("https://github.com/jschicht/RawCopy/raw/master/RawCopy64.exe")
					os.rename(file_name, file_name + ".exe")
					file_name += ".exe"

					print("\n.PST script downloaded successfully to: \"" + str(file_name) + "\"", file=open(abs_logs_path, "a"))

					src = src.replace("\\", "/")
					dst = dst.replace("\\", "/")

					srcArr = src.split("/")
					dstArr = dst.split("/")

					for i in range(len(srcArr)):
						if " " in srcArr[i]:
							srcArr[i] = "\"" + srcArr[i] + "\""
					src = '\\'.join(srcArr)

					for i in range(len(dstArr)):
						if " " in dstArr[i]:
							dstArr[i] = "\"" + dstArr[i] + "\""
					dst = '\\'.join(dstArr)

					dst = dst[:dst.rindex('\\')]

					# locks.insert(len(locks), "lock")
					call(file_name + r' /FileNamePath:' + src + ' /OutputPath:' + dst, stdout=DEVNULL, stderr=STDOUT)
					# lock.pop()

					print("\nExecuted: \"" + file_name + r' /FileNamePath:' + src + ' /OutputPath:' + dst + "\"", file=open(abs_logs_path, "a"))
				except urllib.error.URLError:
					print("\n--> Error getting the .PST script!", file=open(abs_logs_path, "a"))
					if len(sys.argv) == 1:
						print("\n--> Error getting the .PST script! Check your connection please.")
		except:
			print("\n--> Error in copying: \"" + str(src) + "\"", file=open(abs_logs_path, "a"))
			if len(sys.argv) == 1:
				print("\n--> Error in copying: \"" + str(src) + "\"")
			if os.path.exists(dst):
				try:
					os.rmdir(dst)
				except OSError:
					pass


def copytree(src, dst):
	global abs_logs_path
	# global locks

	if not os.path.exists(dst):
		os.makedirs(dst)
	for item in os.listdir(src):
		s = os.path.join(src, item)
		d = os.path.join(dst, item)
		
		if os.path.isdir(s):
			try:
				copytree(s, d)
			except PermissionError as ex:
				os.rmdir(d)
				print("\n--> Error in copying (Permission Denied): \"" + str(d) + "\"", file=open(abs_logs_path, "a"))
				if len(sys.argv) == 1:
					print("\n--> Error in copying (Permission Denied): \"" + str(d) + "\"")
		else:
			try:
				if not s.endswith(".pst"):
					print("\nCopied: \"" + str(s) + "\"", file=open(abs_logs_path, "a"))
					copy2(s, d)
				else:
					try:
						file_name, headers = urllib.request.urlretrieve("https://github.com/jschicht/RawCopy/raw/master/RawCopy64.exe")
						os.rename(file_name, file_name + ".exe")
						file_name += ".exe"

						print("\n.PST script downloaded successfully to: \"" + str(file_name) + "\"", file=open(abs_logs_path, "a"))

						s = s.replace("\\", "/")
						d = d.replace("\\", "/")

						srcArr = s.split("/")
						dstArr = d.split("/")

						for i in range(len(srcArr)):
							if " " in srcArr[i]:
								srcArr[i] = "\"" + srcArr[i] + "\""
						s = '\\'.join(srcArr)

						for i in range(len(dstArr)):
							if " " in dstArr[i]:
								dstArr[i] = "\"" + dstArr[i] + "\""
						d = '\\'.join(dstArr)

						d = d[:d.rindex('\\')]

						# locks.insert(len(locks), "lock")
						call(file_name + r' /FileNamePath:' + s + ' /OutputPath:' + d, stdout=DEVNULL, stderr=STDOUT)
						# locks.pop()

						print("\nExecuted: \"" + file_name + r' /FileNamePath:' + s + ' /OutputPath:' + d + "\"", file=open(abs_logs_path, "a"))
					except urllib.error.URLError:
						print("\n--> Error getting the .PST script!", file=open(abs_logs_path, "a"))
						if len(sys.argv) == 1:
							("\n--> Error getting file! Check your connection.")
			except:
				print("\n--> Error in copying: \"" + str(s) + "\"", file=open(abs_logs_path, "a"))
				if len(sys.argv) == 1:
					print("\n--> Error in copying: " + str(s))
				os.remove(d)	


if __name__ == '__main__':
	try:
		Main()
	except KeyboardInterrupt:
		if len(sys.argv) == 1:
			print("\n")