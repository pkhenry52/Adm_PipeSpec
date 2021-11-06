# Adm_PipeSpec
## Philosophy
The intent of this program is to provide a tool that will generate and maintain an industrial quality piping specification
as well as track the need for updates and revisions.  It is not an engineering tool that well specify the material or
component ratings need for the fluid the specification is being applied to.  That step requires input and reviews from the
Materials/Inspection, Process and Piping engineering groups.  However it is designed to help in eliminating error which may 
occur once the material is selected and the component ratings have been set.
## Design
The program has been developed in python 3 using SQLite as the database, other programs and modules, as is this program, 
are OpenSource and can be used by anyone for non-commercial applications.
Python 3 was selected for its easy of use, SQLite was selected because it limits the number of access points which allow 
modifications to the actual database.  There is no appreciable limit to the number of connections allowed to view the data.  
### Parts
It is the intent of this program that only a small group be allowed to change the data due to its sensitive nature and related
safety issues.  For this reason the program is actually developed in two parts.  One part is for the field Users, 
and can be seen at the repository https://github.com/pkhenry52/User_PipeSpec, it has no entry point into the database 
for any use other than viewing and printing.  The second part is used by the Administrator, the repository you are presently viewing, 
this allows for viewing, printing and changes to the database.

 1. The program allows for easy development, maintenance and update control of a complete plant site commodity pipe specification.
 1. It has the capability to have a custom commodity pipe specification code of up to 15 characters.
 1. The specification is build based on material grade, material type, pressure rating (flange), and corrosion allowance, common in the industry.
 1. The commodity code itself can be unique to each specification or duplicated based on a class such as; STM for steam as a commodity code tied to various
 different steam properties, a specific commodity code such as STMH for high pressure steam properties
 1. A large amount of the data used to build a specification is already built into the data and needs only a click to select.
 1. Information available for the development of a specification is already narrowed down based on the specified material type and grade.
  1. There are numerous forms pre-built in HTML format to hand items which require additional sign off approvals, such as Hydro Test Waivers, Material Substitution Requests, Non-conformance Reports etc.
 1. The building of Scope of Works for the end users has been automated to print out all relevant pipe specification information or just selected sections if needed.
 1. There is a system to import both the Material Substitution Requests and Non-conformance Report, both HTML documents into the database.  The intent is so the can be reviewed at a later data and if needed addressed or even to make amends to the piping specification data.
 1. To help control inspection there is a capability to build an inspection travel sheet for each scope of work.
 1. Because it is a database the data can be updated easily unlike a word document or pdf file. The data can be turned out either as access of a single source on a central server or as a file to each required end user.  The administrator can keep a separate control document which can be updated on the go.
 1. EPCM companies can have numerous different pipe specifications all accessible with this program.

 1. The information covered in each commodity pipe specification includes;
       
 - Piping 
 - Fittings (elbows, tees, laterals, reducers etc.)
 - Flanges
 - Orifice Flanges
 - Gasket Packs
 - Fasteners
 - Unions
 - O-lets
 - Groove Clamps
 - Weld Requirements
 - Tubing
 - Branch Chart
 - Inspection Packs
 - Notes for commodities and pipe components
 - Paint Specification
 - Insulation
 - Gate Valve
 - Globe Valve
 - Plug Valve
 - Ball Valve
 - Butterfly Valve
 - Piston Check Valv
 - Swing Check Valve
 - Specials


The program has been developed in Python 3.7 and is cross platform with Windows and Linux systems.      The program code is open source and free to be downloaded and modified as desired, all though this is not the intent for most users.  It allows for testing and trial turn out to the users with sample data.

The data is located in a collection of 116 data tables, all viewable with SQLite Studio which can be downloaded without charge at https://sqlitestudio.pl .  It is not recommend that the tables be accessed this way as damage to the data schema can occur if the user is not familiar with SQLite.  The Pipe Specification gives full control over all aspects of the data with built in safe guards to prevent this inadvertent damage.

If further information is needed a user manual can be seen at https://github.com/pkhenry52/Adm_PipeSpec/blob/master/Docs/PipeSpecification.pdf.

## Installation

To install the program first download the files using the "Code" then "Download file" buttons.  Then unzip the files into a new directory 'Adm_Pipe_Spec'.  Next install PyInstaller using the 'pip install pyinstaller' in your terminal (Windows used the 'cmd' command, in linux use Ctrl+Alt+t to access the terminal).  This should install the PyInstaller program.  In your terminal change to the directory "Adm_Pipe_Spec' where you downloaded the program.  Then execute the following code 'pyinstaller --onefile Adm_PipeSpec.py', without the quotes, in the terminal. In a few minutes this will produce an excutible file for your present operting system.  If the pyinstaller is not found try to locate the file name pyinstaller.py on your system and copy it to the same directory 'Adm_Pipe_Spec' and rerun the terminal code.  If the pyinstaller worked the 'Adm_Pipe_Spec' directory there will be a folder called 'dist', inside it there is an executible file called 'Adm_PipeSpec'. All you need do now is to copy this new file and the folders 'Data', 'Docs', and 'Forms' to a directory of your choice and run the new created executible file Adm_PipeSpec.  You can delete the original directory "Adm_Pipe_Spec".
NOTE: if you create the executible file in a Windows enviroment it will only run in Windows the same applies for Linux systems.
