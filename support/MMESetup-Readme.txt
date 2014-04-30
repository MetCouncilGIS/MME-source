MINNESOTA METADATA EDITOR V3.1.1 15sEP2011

**************************************************************************************************************
Installation and Configuration Instructions for the Minnesota Metadata Editor (MME)
**************************************************************************************************************

MME Installation RequirMMEnts:

  -> The MME may be run as a standalone product or as an ArcGIS Extension. 
  
  -> In order to install MME and run in either mode, Microsoft .NET Framework 3.5 must be installed on your machine.You can determine if the .NET Framework is on your machine by going to Start->Control Panel->Add or Remove Programs.  
You should see Microsoft .NET Framework 3.5 listed.  Multiple versions of Microsoft .Net Framework (e.g., 1,2,3) may exist on your machine simultaneously. Microsoft's .NET Framework 3.5 is freely available and can be accessed from Microsoft's website.  

http://www.microsoft.com/download/en/details.aspx?displaylang=en&id=21


  -> If you plan to use the MME as an ArcGIS Extension, then you must also have the following: 
    -ArcGIS 10 or higher must be installed on your machine.  
    -The ArcGIS 10.0 (Desktop) FGDC Metadata Style Patch must be installed on your machine. This patch is free and is           available at:
http://resources.arcgis.com/content/patches-and-service-packs?fa=viewPatch&PID=160&MetaID=1637


  -> The application is data-driven using a Microsoft Access database, metadata.mdb. The primary reason for editing the database is to assign default values that apply to particular agency or group. These default values may be set for each individual. You will require Microsoft Access in order to edit the database. Once edited, it can, however, be distributed to other machines that do not have Access installed. The application, itself, does not require Access in order to run.
   
**************************************************************************************************************
Installing the Minnesota Metadata Editor
**************************************************************************************************************
1. Double click the MME set up file (MME_3.1.1.msi).
2. Follow set-up instructions.
3. The default installation location is C:\Program Files\MnGeo\Minnesota Metadata Editor\.


**************************************************************************************************************
Accessing the EPA Metadata Editor
**************************************************************************************************************
MME may be accessed in Standalone Mode by going to Start -> Programs -> Minnesota Metadata Editor.

MME may be accessed as an ArcGIS Extension by taking the following steps: 
1. Open ArcCatalog and navigate to Customize -> Toolbars. 
2. Select the MGMG Metadata Toolbar. 
3. Once the MGMG Metadata Toolbar is open, navigate to your directory or database of choice in the ArcCatalog table of contents. 
4. Select your data set or metadata record in the 'Contents' window in ArcCatalog. 
5. Click on the 'Edit MGMG Metadata' button in the user interface. 