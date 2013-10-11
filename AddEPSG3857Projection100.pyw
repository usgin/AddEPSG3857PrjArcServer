'''
Edit Arc Config File for Arc 10.0

Edit the Arc Server Config File adding <ListSupportedCRS>EPSG:3857</ListSupportedCRS> 
to the node <Extension> <TypeName>WMSServer</TypeName> <Properties>
Created on Sept 11, 2013
@author: Jessica Good Alisdairi, AZGS jessica.alisdairi@azgs.az.gov 
'''

import os, glob, subprocess
from xml.dom.minidom import parse, parseString
import Tkinter, tkFileDialog
from Tkinter import *


# Main function for the Well Log Data Converter Tool
def main(argv=None):
    root = Tkinter.Tk()
    root.title("Edit Arc Config File")
    root.minsize(0, 0)
    # Create the frame for the conversion option buttons
    labelframe1 = LabelFrame(root, text = "Add <ListSupportedCRS>EPSG:3857</ListSupportedCRS> to Arc Config File")
    Label(labelframe1,text='Usual Config File Location: <ArcGIS install location>\Server\user\cfg').pack(pady=1)
    labelframe1.pack(fill = "both", expand = "yes")
    # Create a frame for both buttons
    bframe = Frame(labelframe1)
    bframe.pack()
    # Create a frame for the 1st button
    b1frame = Frame(bframe, bd = 5)
    b1frame.pack(side = LEFT)
    # Create the 1st button
    B = Tkinter.Button(b1frame, text = "Edit Single Config File", command = EditSingleFile)
    B.pack()
    # Create a frame for the 2nd button
    b2frame = Frame(bframe, bd = 5)
    b2frame.pack(side = LEFT)
    # Create the 2nd button
    B2 = Tkinter.Button(b2frame, text = "Edit all Config Files in a Folder", command = EditFolder)
    B2.pack()
    
    # Create the frame for the messages list
    global textFrame
    msgFrame = LabelFrame(root, text="Messages")
    # Create the frame for the text
    textFrame = Tkinter.Text(msgFrame, height = 30)
    # Create the scrollbar
    yscrollbar = Scrollbar(msgFrame)
    yscrollbar.pack(side = RIGHT, fill = Y)
    textFrame.pack(side = LEFT, fill = Y)
    yscrollbar.config(command = textFrame.yview)
    textFrame.config(yscrollcommand = yscrollbar.set)
    msgFrame.pack(fill = X, expand = "yes")

    # Create the frame for the Exit button
    b3frame = Frame(root, bd = 5)
    b3frame.pack()
    # Create the Exit button
    B3 = Tkinter.Button(b3frame, text = "Exit", command = ExitEdit)
    B3.pack()
    
    root.mainloop()
    return
   
# Prompt for file to convert
def EditSingleFile():
    # Clear the text frame
    textFrame.delete("1.0", END)
    
    cfgFile = tkFileDialog.askopenfilename(filetypes=[("Arc Config Files","*.cfg")])
    # If cancel was pressed
    if cfgFile == "":
        return
  
    cfgFiles = []
    cfgFiles.append(cfgFile)
    Edit(cfgFiles)
    return

# Prompt for folder which holds all the Excel files to convert 
def EditFolder():
    # Clear the text frame
    textFrame.delete("1.0", END)
    
    global cfgBaseFile
    path = tkFileDialog.askdirectory()
    # If cancel was pressed
    if path == "":
        return

    # Find all Config files in the folder
    cfgFiles = glob.glob(path + "/*.cfg")
    
    Edit(cfgFiles)
    return

# Start the editing
def Edit(cfgFiles):
    
    # Read the cfg files
    for cfgFile in cfgFiles:
        cfgBaseFile = os.path.basename(cfgFile)
        try:
            # Open the file for reading
            f = open(cfgFile, 'r')
            cfg = f.read()
            domCFG = parseString(cfg)
        except:
            return
    
        # See if the ListSupportedCRS node already exists
        nodeListSupportedCRS = domCFG.getElementsByTagName("ListSupportedCRS")
        
        # If the ListSupportedCRS node already exists don't edit the file
        if nodeListSupportedCRS.length != 0:
            Message("Already Edited " + cfgBaseFile)
       
        # If the ListSupportedCRS node doesn't exist edit the file
        else:
            # Get the Extension nodes
            for nodeExtension in domCFG.getElementsByTagName("Extension"):
                # Get the TypeName nodes within the Extension nodes
                for nodeTypeName in nodeExtension.getElementsByTagName("TypeName"):
                    # If the value of the TypeName node is WMSServer
                    if nodeTypeName.childNodes[0].nodeValue == "WMSServer":
                        # Get the Properties node within the Extension node
                        for nodeProperties in nodeExtension.getElementsByTagName("Properties"):
                            # Create node <ListSupportedCRS>EPSG:3857</ListSupportedCRS>
                            x = domCFG.createElement("ListSupportedCRS")
                            txt = domCFG.createTextNode("EPSG:3857")
                            x.appendChild(txt)
                            # Append the new node
                            nodeProperties.appendChild(x)
            
            f.close()
            # Open the cfg for writing
            f = open(cfgFile, 'w')
            f.write(domCFG.toxml())
            f.close()
            
            Message("Edited " + cfgBaseFile)   
    
            #Restart Service to refresh - Contributed by Thomas O'Malley (thms.omalley@gmail.com) at ISGS
            try:
                subprocess.call('net stop "ArcGIS Server Object Manager"')
                subprocess.call('net start "ArcGIS Server Object Manager"')
            except:
                print "ArcGIS SOM Restart Failed"
            
# Exit the program
def ExitEdit():
    exit()
    return

# Message Box
def Message(message):
    textFrame.insert(END, message + "\n")
    return

if __name__ == "__main__":
    main()