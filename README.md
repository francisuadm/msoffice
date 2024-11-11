# msoffice
All related with Microsoft Office


````
Get-WmiObject Win32_CodecFile | Select-Object Name, Description, Version
# or
Get-WmiObject Win32_CodecFile | FT -Property Name, Description, Version, Group, Manufacturer
````

### Verifying Codec Installation

Here's how you can check if these specific codecs are installed:

#### H.264 Codec:

Open a Command Prompt window and type:

````
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" /s /f "H.264"
````
This will search the registry for any mention of the H.264 codec.


#### AAC Codec:

Open a Command Prompt window and type:

````
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" /s /f "AAC"
````
This will search the registry for any mention of the AAC codec.

#### MP3 Codec:

Check for the Fraunhofer IIS MPEG Layer-3 Codec in your output:

C:\WINDOWS\system32\L3CODECA.ACM          Fraunhofer IIS MPEG Layer-3 Codec 1.9.0.401

It seems like you already have the MP3 codec installed.



## Here's how you can convert your media files using VLC Media Player to ensure they work with PowerPoint:

````
Download and Install VLC Media Player: If you haven't already, download VLC Media Player from its official website and install it on your device.

Open VLC Media Player: Launch VLC Media Player once it's installed.

Convert/Save Media Files: Go to Media > Convert/Save. A new window will open2.

Add Media File: Click on Add to select the media file you want to convert.

Choose Output Format: In the Profile dropdown, select MP4 for video files or M4A for audio files.

Convert/Save: Click on Convert/Save to start the conversion process.

Once the conversion is complete, your media files should be in a format compatible with PowerPoint. If you encounter any issues or need further assistance, feel free to ask! 😊
````






I'm sorry to hear you're having trouble with your PowerPoint presentation! Let's try a few steps to resolve this issue:

Check Media Compatibility: Go to the File tab, select Info, and click on Optimize Media Compatibility. This will help ensure that your media files are compatible with other devices and versions of PowerPoint2.

Embed Media Files: If your media files are linked rather than embedded, PowerPoint might not be able to find them. Go to File > Info > View Links and embed the media files by selecting Break Link.

Convert Media Files: Ensure that your media files are in a format supported by PowerPoint. You can convert them to a compatible format using tools like VLC or online converters4.

Update Codecs: Sometimes, missing codecs can cause playback issues. Download and install a codec pack like K-Lite Codec Pack Basic5.

Repair Corrupted Files: If your media files are corrupted, you can use tools like 4DDiG File Repair to fix them.

Run PowerPoint in Safe Mode: Close PowerPoint, press Win + R, type powerpnt /safe, and press OK. This can help identify if add-ins or other settings are causing the issue6.

If none of these steps work, please let me know, and we can try some other troubleshooting methods. Have you tried any of these steps already
