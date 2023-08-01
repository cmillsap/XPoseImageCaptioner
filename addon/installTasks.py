import urllib.request
import shutil 
from logHandler import log
import traceback
import globalVars
import os
import ui 


def copyURLtoFile(url, tempFilename):
	req = urllib.request.Request(url)
	req.add_header('User-Agent', 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36')
	with urllib.request.urlopen(req) as response:
		with open(tempFilename, "wb") as f:
			# Copy the binary content of the response to the file
			shutil.copyfileobj(response, f)	


def onInstall():
	#build a path to the addon directory and weights path 
	weightsPath = os.path.join(globalVars.appArgs.configPath, 'addons\\\XPoseImage Captioner.pendingInstall\\globalPlugins\\dist\\weights\\model_base_caption_capfilt_large.pth')
	try:
		copyURLtoFile('https://storage.googleapis.com/sfr-vision-language-research/BLIP/models/model_base_caption_capfilt_large.pth',weightsPath)
	except Exception as e: 
		log.error(f"Exception = {e}")
		traceback.print_exception(type(e), e, e.__traceback__)
		ui.message("BLIP Neural Network wieghts could not be downloaded. Installation can not proceed. ")
