# SESE - SolidWorks Easy STL Export
SESE is a SolidWorks Addin compatible with SolidWorks 2011 SP5 and newer.

It provides an easy way to export the active document to an STL file used for 3D printing.

## Usage
After registering the DLL you may have to activate it in SolidWorks, maybe not.
Then you will have an extra menu entry in *Tools* which is called *Export STLs*.
This will export the active document to STL to the same location and the same filename.
If the active document is a multi body part then each body is exported to an own STL file named after the body's name.
You can rename a body in the feature tree of Solidworks.
In addition if you have an extra coordinate system created which its name is *STL* this will be the coordinate system used for the STL export.