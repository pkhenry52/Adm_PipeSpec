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

Additional information can be seen https://github.com/pkhenry52/Adm_PipeSpec/blob/master/Docs/PipeSpecification.pdf
