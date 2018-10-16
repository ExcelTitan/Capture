# Capture
With this Add-in all of your saved colours are available to all workbooks! 
Just open the Add-in, select your saved colour and click ‘+ to recent colours’.

## Compatibility: 
- PC Only
- Excel 2007 - 2016
- Works on 32-bit or 64-bit

## So Easy to Use:
Capture any colour on the screen (even outside the Excel app) by clicking on the colour picker, drag and release on the
colour anywhere on the screen to capture that colour. The RGB, Hex and Long value of the captured colour will populate
in the ‘Convert Colours’ section.
From here you can add the captured colour to your Recent Colours Panel, save the colour to your log and even convert the colour to the RGB, Hex or Long Values

### Definitions:
##### RGB:
The RGB color model is an additive color model in which red, green and blue light are added together in various ways to reproduce a broad array of colors. The first value pair refers to red, the second to green and the third to blue, with decimal values ranging from 0 to 255.  
##### Hex:
A color hex code is a way of specifying color using hexadecimal values. The code itself is a hex triplet, which represents three separate values that specify the levels of the component colors. The code starts with a pound sign (#) and is followed by six hex values or three hex value pairs (for example, #AFD645). The code is generally associated with HTML and websites, viewed on a screen, and as such the hex value pairs refer to the RGB color space.
#### Long:
The long that is returned from the Interior.Color property is a decimal conversion of the typical hexidecimal numbers that we are used to seeing for colors in html e.g. Yellow in Hex is "FFFF00" but the Excel long value is 65535. 


----
## Installation
- Unzip the zip file to extract the contents
- Move the add-in to your local Microsoft AddIns directory.
- Open Excel
- Click the File tab
- Click Options
- Click the Add-Ins category on the left bar of the Excel Options Dialog Box
- In the Manage box near the bottom, click Excel Add-ins, and then click Go. The Add-Ins dialog box appears.
- Click Browse
- Navigate to the xla or xlam add-in and select it.
- Make sure the checkbox next to the add-in is checked, then click OK.

If the add-in keeps disappearing when you restart Excel, follow the directions below.  
***Important Note:*** There are a few additional steps to this installation due to an Office Security Update released in July 2016.


#### Trust the Folder Location
The folder that the add-in file is saved in needs to be added as a Trusted Location in Excel. The instructions on how to trust the folder location are below.  
***Open the Excel Options menu:***
- File > Options
- Open the Trust Center menu and Add a new location
- Trust Center > Trust Center Settings… > Trusted Locations > Add new location
- Browse for the folder that you saved the add-in file in.
- Press OK, the folder should now appear in the Trusted Locations list.
- Press OK on the Trust Center and Excel Options menu to close the menus.  
 
The add-in is now in a trusted location and the add-in will appear every time you open Excel.  


#### Unblock the File
Some users still have issues with the add-in’s ribbon disappearing after trusting the folder location. If this happens to you, you will need to Unblock the file by changing a file property.  

- Locate the Add-in file (.xla, .xlam) in Windows Explorer.
- Right-click the file and select Properties.
- At the bottom of the General tab you should see a Security section. Check the box that says Unblock.
- Press the OK button.

Close Excel completely and re-open it. The add-in should now load and any custom ribbons will appear.  

You will only need to do this unblock one time. However, if you download an updated version of the file then you will have to repeat the steps above to unblock it.  
 
Hopefully these additional security steps will be fixed in a future update to Office. If you completed all the steps above then you should see the add-ins ribbon tab load every time you open Excel.
