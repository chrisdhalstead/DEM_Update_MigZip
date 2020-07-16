# DEM_Update_MigZip
This script is intended for use when migrating from View Composer Persistent Disks to FSLogix Profile Containers. As part of a broader migration scenario, DEM migration.zip files are automatically enumerated, unzipped, edited, and zipped.

### Instructions

1. Save the VBS to a location accessible from an admin PC.

2. Edit the variables in this section to align with your environment.<br/>

   ![dem1](https://github.com/chrisdhalstead/DEM_Update_MigZip/blob/master/Images/dem1.jpg)

3. Run the script from an admin PC.<br/>

4. For each migration.zip found in the DEM profile archives path, Flex Profiles.reg will extracted to a local working directory and updated.<br/>
   ![dem2](https://github.com/chrisdhalstead/DEM_Update_MigZip/blob/master/Images/dem2.jpg)

5. A log file is appended for each profile archive that is found and edited.<br/>

   ![dem3](https://github.com/chrisdhalstead/DEM_Update_MigZip/blob/master/Images/dem3.jpg)

