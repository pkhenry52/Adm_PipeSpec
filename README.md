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

 1.   1. The program allows for easy development, maintenance and update control of a complete plant site commodity pipe specification.
 1.   2. It has the capability to have a custom commodity pipe specification code of up to 15 characters.  The code outlined for the development of a specification does not have to be that shown on the final documentation.
    3. The specification is build based on material grade, material type, pressure rating (flange), and corrosion allowance, common in the industry.
    4. The commodity code itself can be unique to each specification or duplicated based on a class such as STM for steam.  This possible because each commodity code is built around a specific set of properties; pressure, temperature, fluid category etc.
    5. A large amount of the data used to build a specification is already built into the data and needs only a click to select.
    6. Information available for the development of a specification is already narrowed down based on the specified material type and grade.
    7. The program is actually two separate parts, one is for the administrator the other for the end users.  This limits the possibility of unauthorized changes to the data.
    8. The data itself is based on SQLite one of the most used database engines in the world and is opensource meaning licensing is not required, https://sqlite.org/index.html .
    9. SQLite limits the number of sign ins to a database for updating and modifications, it does not limit the number of viewers, in this case your end users.
    10. There are numerous forms pre-built in HTML format to hand items which require additional sing off approvals, such as Hydro Test Waivers, Material Substitution Requests, Non-conformance Reports etc.
    11. The building of Scope of Works for the end users has been automated to print out all relevant pipe specification information or just selected sections if needed.
    12. The administrator has  a single screen location to update the support tables as needed, and build the complete commodity pipe specification.
    13. There is a system to import both the Material Substitution Requests and Non-conformance Report, both HTML documents into the database.  The intent is so the can be reviewed at a later data and if needed addressed or even to make amends to the piping specification data.
    14. To help control inspection there is a capability to build an inspection travel sheet for each scope of work.
    15. Because it is a database the data can be updated easily unlike a word document or pdf file. The data can be turned out either as access of a single source on a central server or as a file to each required end user.  The administrator can keep a separate control document which can be updated on the go.
    16. EPCM companies can have numerous different pipe specifications all accessible with this program.

    17. The information covered in each commodity pipe specification includes;

                        ▪ Piping
                        ▪ Fittings (elbows, tees, laterals, reducers etc.)
                        ▪ Flanges
                        ▪ Orifice Flanges
                        ▪ Gasket Packs
                        ▪ Fasteners
                        ▪ Unions
                        ▪ O-lets
                        ▪ Groove Clamps
                        ▪ Weld Requirements
                        ▪ Tubing
                        ▪ Branch Chart
                        ▪ Notes for commodities and pipe components.
                        ▪ Inspection Packs
                        ▪ Paint Specification
                        ▪ Insulation
                        ▪ Gate Valve
                        ▪ Globe Valve
                        ▪ Plug Valve
                        ▪ Ball Valve
                        ▪ Butterfly Valve
                        ▪ Piston Check Valve
                        ▪ Swing Check Valve
                        ▪ Specials


The program has been developed in Python 3.7 and is cross platform with Windows and Linux systems.      The program code is open source and free to be downloaded and modified as desired, all though this is not the intent for most users.  It allows for testing and trial turn out to the users with sample data.

The data is located in a collection of 116 data tables, all viewable with SQLite Studio which can be downloaded without charge at https://sqlitestudio.pl .  It is not recommend that the tables be accessed this way as damage to the data schema can occur if the user is not familiar with SQLite.  The Pipe Specification gives full control over all aspects of the data with built in safe guards to prevent this inadvertent damage.

If further information is needed a user manual can be seen at https://github.com/pkhenry52/Adm_PipeSpec/blob/master/Docs/PipeSpecification.pdf.
