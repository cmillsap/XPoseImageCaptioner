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
import urllib.request
from urllib.parse import urlparse 
import shutil 
import struct
import time
import tempfile
import os
import wx 
from logHandler import log
import traceback
from comtypes.client import CreateObject as COMCreate
from comtypes.gen.ISimpleDOM import ISimpleDOMDocument
import controlTypes
import subprocess
import addonHandler

ADDON_SUMMARY = addonHandler.getCodeAddon().manifest["summary"]


class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	@script(
		# Translators: message presented in input help mode, when user presses the shortcut keys for this addon.
		description=_('Describe an image with {name}').format(name=ADDON_SUMMARY),
		gesture="kb:NVDA+x",
	)
	def script_captionPhotograph(self, gesture):
		wx.InitAllImageHandlers()
		filename = self.get_selected_file()
		if (filename is not False):
			self.captionImageFile(filename)
		else: 
			# try to obtain information about the in-focus object
			currentObject = api.getNavigatorObject()
			if (currentObject is not None): 
				name = currentObject._get_name()
				roleName = currentObject._get_roleText()
				log.info(f'Not explorer, name is {name} with role {roleName}')
				log.info(f'object details {currentObject}')
				finalAttributes = currentObject._get_IA2Attributes()
				if 'tag' in finalAttributes:
					if 'src' in finalAttributes: 
						imagefilename = finalAttributes['src']
						log.info(f'image name is {imagefilename}')
						urlComponents = urlparse(imagefilename)
						if (urlComponents.netloc == ""):
							navUrl = self.get_url_from_nav_object()
							navUrlcomponents = urlparse(navUrl)
							extension =  urlComponents.path[-4:].lower()
							if (extension.endswith(self.image_extensions) == True):
								finalUrl = navUrlcomponents._replace(path = urlComponents.path).geturl() 
								log.info(f'final url = {finalUrl}')
								tmpFile = self.getTempFileName(finalUrl)
								self.captionImageURL(finalUrl, tmpFile)
								return 
							else: 
								ui.message('Image file type is not readable. ')
						else: 
							urlImgFile = urlComponents.path 
							extension = urlImgFile[-4:].lower()
							log.info(f'file name is {urlImgFile}')
							serverpath, imgFile = os.path.split(urlImgFile)
							log.info(f'simple file name is {imgFile}')
							tempFile = tempfile.gettempdir() + '\\' + imgFile
							log.info(f'temp file name is {tempFile}')
							if (extension.endswith(self.image_extensions) == True):
								# valid image. see if image captioning is possible. 
								if (urlComponents.netloc != ""): 
									# url contains full text so not relative. Load and send to cpationing 
									netloc = urlComponents.netloc
									log.info(f'netloc = {netloc}')
									log.info(f'image file name {imagefilename}')
									self.captionImageURL(imagefilename, tempFile)
									return 			
					else: 
						for key, value in finalAttributes.items():
							log.info(key+" "+ str(value))
							url2 = self.get_url_from_nav_object()
						log.info(f'src property not present. ALT tag in Firefox {url2}')
						navObj = api.getNavigatorObject()

				else:
					log.info('img property not present. ')
			
    
	def getTempFileName(self, imgUrl):
		urlComponents = urlparse(imgUrl)
		urlImgFile = urlComponents.path 
		log.info(f'file name is {urlImgFile}')
		serverpath, imgFile = os.path.split(urlImgFile)
		log.info(f'simple file name is {imgFile}')
		tempFile = tempfile.gettempdir() + '\\' + imgFile
		return tempFile
	
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
		self.addonPath = os.path.join(globalVars.appArgs.configPath, "addons\\\XPoseImage Captioner\\globalPlugins")
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

	def captionImageFile(self, imgFileName): 
		theImg = wx.Image()
		extension = imgFileName[-4:].lower()
		if (extension.endswith(self.image_extensions) == True):
			try:
				theImg = wx.Image(imgFileName)
				self.captionImage(theImg)
			except: 
				ui.message("Image file could not be read. ")
				return 
		else: 
			ui.message(f'{imgFileName} is not a recognized image type')
		
	def captionImage(self, theImg): 
		imgData = theImg.GetData()
		imgBytes = bytes(imgData)
		width = theImg.GetWidth()
		height = theImg.GetHeight()
		time.sleep(0.1)
		server_status = self.get_server_status(self.commandMap)
		if (server_status == 0):
			ui.message('Captioning Server is not yet loaded. ')
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
			caption = caption.decode("utf-8")
			self.send_response_from_client(self.commandMap, 0)#client has no more traffic 
			self.send_response_from_server(self.commandMap, 0)
			log.info(f'The image size is {width} by {height} and uses {len(imgBytes)} bytes of memory. Caption: {caption}')
			ui.browseableMessage(f"caption : {caption}", title='Captioned image', isHtml=False)
	
	def captionImageURL(self, url, tempFileName):
		try:
			req = urllib.request.Request(url)
			req.add_header('User-Agent', 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36')
			with urllib.request.urlopen(req) as response:
				with open(tempFileName, "wb") as f:
					 # Copy the binary content of the response to the file
					shutil.copyfileobj(response, f)
			self.captionImageFile(tempFileName)
		except Exception as e: 
			ui.message("Image file could not be read. ")
			log.error(f"url = {url}. Exception = {e}")
			traceback.print_exception(type(e), e, e.__traceback__)
			return 
	
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
	
	def get_URL_from_object(self, startObject):
		searchObject = startObject 
		while searchObject.role != controlTypes.Role.DOCUMENT:
			searchObject = searchObject.parent
		try:
			doc =searchObject.IAccessibleObject.QueryInterface(ISimpleDOMDocument)
			return doc.URL
		except:
			return ""
		
	def get_url_from_nav_object(self): 
		URL = None 
		obj = api.getNavigatorObject()
		try:
			URL = obj.treeInterceptor.documentConstantIdentifier
			log.info("Masked  : " + URL)
			URL = urllib.parse.unquote(URL)
			log.info("Unmasked: " + URL)
		except:
			log.info("URL not found in get_url_from_nav_object")
			return None 
		return URL 
