# XPoseImage Captioner for NVDA #
----------
## Overview 
This addon allows for the AI captioning of JPEG and PNG images from Windows File Explorer,
Microsoft Edge, Google Chrome, and Firefox. It does all inference on the user's local machine, rather than relying on a remote server to caption the image. 
To use, first select an image with the navigation cursor, then press NVDA+x to activate the addon. The addon will say "Captioning, please wait" as the AI captions the image. Thia process will take from one to five seconds depending on your machine's CPU speed. A window will open after the AI completes the caption showing the text of the caption and the caption text will be read. You can dismiss the caption window by pressing Escape. 

## Getting the most out of XPoseImage Captioner
There are several things to be aware of when using the XPoseImageCaptioner to get the best results: 

1. AI captioning works best for photographs and cartoons or other artwork. It also can work fairly well for memes and ads. It does not work well for charts and is not a replacement for OCR. If you have an image of a text document, use an OCR addon rather than XPoseImageCaptioner. 
2. AI captioning can tell you what is in an image but can't tell you why it is there. ALT text should still be used to find out about an image's context. For example, on a news site you may see an image with the ALT text "a general gives testimony in a congressional hearing about the millitary budget", the AI caption could be something like "a man in a formal millitary uniform speaks into a microphone while seated in a wood paneled room". The AI caption tells you what is in the image, but the ALT text ideally should tell you why its there. 
3. The BLIP neural network, on which the XPoseImageCaptioner addon is based, can only output English text. Retraining the model to support languages other than English is not feasable at this time. 
4. While the captions produced are currently very close to the state of the art for AI captioning of images, they are not always 100% accurate. Please use with caution and common sense and never in place of OCR. Also, do not rely on the output for dangerous or high-risk situations. 
5. Currently, XPoseImageCaptioner works for websites that don't require a login. For example, the public pages of organizations such as [Guiding Eyes for the Blind](https://www.guidingeyes.org/) or [CNN](https://www.cnn.com/). Pages that require a login, such as Facebook or Twitter, are not yet supported because the addon must download an image from the web site independantly to caption it and can't do so if a login is required. As a workaround, any image from sites requiring a login could be downloaded to the local machine and captioned using the addon in File Explorer. 
6. Full Firefox support for XPoseImageCaptioner requires FireFox version 116 or later. 



Incorporates the BLIP neural network and supporting code from the paper 
__BLIP: Bootstrapping Language-Image Pre-training for Unified Vision-Language Understanding and Generation__ by Junan Li et al (https://arxiv.org/abs/2201.12086)


XPoseImageCaptioner &copy; 2024 Christopher Millsap, All Rights Reserved. Distributed under the **BSD 3 Clause License**. This addon is not affiliated with or endorsed by Salesforce.com 

Uses the BLIP neural network and supporting code. &copy; 2024 SalesForce.com, All Rights Reserved. Distributed under the **BSD 3 Clause License**. 

See License.txt for full details. 
