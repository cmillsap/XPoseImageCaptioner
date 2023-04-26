# Photo Captioner add-on for NVDA using the BLIP model 
# Model by Junnan Li at el in paper BLIP: Bootstrapping Language-Image Pre-training for Unified Vision-Language Understanding and Generation  https://arxiv.org/abs/2201.12086
# NVDA addon by Christopher Millsap 
# Much thanks to the NVDA NAO project for examples in interacting with file explorer

import globalPluginHandler
from scriptHandler import script
import ui
import api
import mmap
import globalVars
import struct
import time
import os
import wx 
from logHandler import log
from comtypes.client import CreateObject as COMCreate
import subprocess


class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	@script(gesture="kb:NVDA+shift+v")
	def script_captionPhotograph(self, gesture):
		wx.InitAllImageHandlers()
		filename = self.get_selected_file()
		if (filename is not False):
			extension = filename[-4:].lower()

			theImg = wx.Image()
			if (extension.endswith(self.image_extensions) == True):
				try:
					theImg = wx.Image(filename)
				except: 
					ui.message("Image file could not be read. ")
					return 
				imgData = theImg.GetData()
				imgBytes = bytes(imgData)
				width = theImg.GetWidth()
				height = theImg.GetHeight()
				time.sleep(0.1)
				server_status = self.get_server_status(self.commandMap)
				if (server_status == 0):
					ui.message(f'Captioning Server is not yet loaded. ')
				else: 
					ui.message("Captioning, please wait...")
					self.commandMap.seek(self.image_width_offset * 4) 
					self.commandMap.write(width.to_bytes(4, byteorder='big')) 
					self.commandMap.write(height.to_bytes(4, byteorder='big'))
					mm = mmap.mmap(-1, len(imgBytes) + 4, "BLIPImage")
					mm.write(imgBytes)
					mm.flush()
					self.send_response_from_client(self.commandMap, 77) #all image information is available 
					self.await_response_from_server(self.commandMap, 78) #captioning complete
					self.commandMap.seek(self.commandSize)
					caption = self.commandMap.readline()
					mm.close()
					self.send_response_from_client(self.commandMap, 0)#client has no more traffic 
					self.send_response_from_server(self.commandMap, 0)
					log.info(f'The image size is {width} by {height} and uses {len(imgBytes)} bytes of memory. Caption: {caption}')
					ui.browseableMessage(f"caption : {caption}", isHtml=False)
			else: 
				ui.message(f'{filename} is not a recognized image type')
		else: 
			# try to obtain information about the in-focus object
			currentObject = api.getFocusObject()
			if (currentObject is not None): 
				name = currentObject._get_name()
				roleName = currentObject._get_roleText()
				# permits getting graphics coordinates from current object. 
				# currLocation = currentObject._get_treeInterceptor().currentFocusableNVDAObject.location
				# theBitmap = screenBitmap.ScreenBitmap(384,384)
				# buffer = theBitmap.captureImage(currLocation.left, currLocation.top,(currLocation.right - currLocation.left), (currLocation.bottom - currLocation.top))
				# theImg = wx.Image(384,384,buffer)
				# theImg.SaveFile("C:\\BLIPDemo\\testfromscr.bmp")
				# print(f'currentObject location  {currLocation.left} {currLocation.top} {currLocation.right} {currLocation.bottom}')
				# print(f'{dir(buffer)}')
				ui.message(f'Not explorer, name is {name} with role {roleName}')
			ui.message("No image selected. ")

	def __init__(self):
		super(GlobalPlugin, self).__init__()
		self.commandSize = struct.calcsize(">iiiii")
		self.commandMap = mmap.mmap(-1, self.commandSize + 256, "BLIPCommand")
		self.server_message_offset = 0 
		self.client_message_offset = 1 
		self.server_status_offset = 2
		self.image_width_offset = 3 
		self.image_height_offset = 4 
		self.scratchpadPath = "C:\\Users\\chris\\AppData\\Roaming\\nvda\\scratchpad\\globalPlugins"
		self.usingScratchpad = False
		self.addonPath = os.path.join(globalVars.appArgs.configPath, "addons\\\imageCaptioner\\globalPlugins")
		self.server_not_ready = 0
		self.server_ready = 40 
		self.image_extensions = ("jpg", "jpeg", "png", "gif", "bmp")
		currentWorkingDirectory = os.getcwd()
		if self.usingScratchpad: 
			self.server_location = self.scratchpadPath + "\\dist\\"
		else: 
			self.server_location = self.addonPath + "\\dist\\"
		os.chdir(self.server_location)
		log.info(f"launching Image Captioning Server at {self.server_location}") 
		startupinfo = subprocess.STARTUPINFO()
		startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
		self.server = subprocess.Popen("CaptionServer.exe", startupinfo=startupinfo)
		os.chdir(currentWorkingDirectory)

	def terminate(self):
		self.send_response_from_client(self.commandMap, 23) 
		log.info("Image Captioning server shutting down. ")
		# shut down server if NVDA shuts down. 
	
	def await_response_from_client(self, map, validClientBytes): 
		readVal = 0
		foundValue = False
		while (foundValue is False):
			time.sleep(0.01)
			readVal = map[self.client_message_offset]
			count = validClientBytes.count(readVal); 
			if (count > 0): 
				foundValue = True 
				log.info(f'received {readVal}')
				return readVal    
	
	def send_response_from_server(self, map, commandByte):
		map[self.server_message_offset] = commandByte 
		log.info(f'sent {commandByte}')
	
	def set_server_status(self, map, statusValue): 
		map[self.server_status_offset] = statusValue 

	def get_server_status(self, map): 
		return map[self.server_status_offset]  

	def await_response_from_server(self, map, commandByte): 
		readVal = 0
		while (readVal != commandByte):
			time.sleep(0.01)
			readVal = map[self.server_message_offset]
		return readVal  

	def send_response_from_client(self, map, commandByte):
		map[self.client_message_offset] = commandByte 
		log.info(f'sent {commandByte}')

	def get_selected_file_explorer(self, obj=None):
		if obj is None: 
			obj = api.getForegroundObject()
		file_path = False
		# We check if we are in the Windows Explorer.
		if self.is_explorer(obj):
			desktop = False
			try:
				global _shell
				if not _shell:
					_shell = COMCreate("shell.application")
				# We go through the list of open Windows Explorers to find the one that has the focus.
				for window in _shell.Windows():
					if window.hwnd == obj.windowHandle:
						# Now that we have the current folder, we can explore the SelectedItems collection.
						file_path = str(window.Document.FocusedItem.path)
						break
				else:  # loop exhausted
					desktop = True
			except:
				try:
					windows = self.get_selected_files_explorer_ps()
					if windows:
						if str(obj.windowHandle) in windows:
							file_path = windows[str(obj.windowHandle)]
						else:
							desktop = True
				except:
					pass
			if desktop:
				desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
				file_path = desktop_path + '\\' + api.getDesktopObject().objectWithFocus().name
				#if not os.path.isfile(file_path) and not os.path.isdir(file_path): file_path = False
		return file_path
	
	def is_explorer(self, obj=None):
		if obj is None: 
			obj = api.getForegroundObject()
		#return obj and (obj.role == api.controlTypes.Role.PANE or obj.role == api.controlTypes.Role.WINDOW) and obj.appModule.appName == "explorer"
		return obj and obj.appModule and obj.appModule.appName and obj.appModule.appName == 'explorer'

	def get_selected_file(self, obj=None):
		file_path = False
		if obj is None: 
			obj = api.getForegroundObject()
		file_path = self.get_selected_file_explorer(obj)
		return file_path

	def get_selected_files_explorer_ps(self):
		si = subprocess.STARTUPINFO()
		si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
		cmd = "$OutputEncoding = [console]::InputEncoding = [console]::OutputEncoding = New-Object System.Text.UTF8Encoding; (New-Object -ComObject 'Shell.Application').Windows() | ForEach-Object { echo \\\"$($_.HWND):$($_.Document.FocusedItem.Path)\\\" }"
		cmd = "powershell.exe \"{}\"".format(cmd)
		try:
			p = subprocess.Popen(cmd, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, startupinfo=si, encoding="utf-8", text=True)
			stdout, stderr = p.communicate()
			if p.returncode == 0 and stdout:
				ret = {}
				lines = stdout.splitlines()
				for line in lines:
					hwnd, name = line.split(':',1)
					ret[str(hwnd)] = name
				return ret
		except:
			pass
		return False
