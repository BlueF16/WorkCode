# Matlab SV200 Data Compiler

These scripts are intended for use, to compile data from station generated CSV files.

## Components

* CSV Parser.m: Reads each stations CSV files seprately and creates and excel sheet for that station with all relevant data used.

* DataCompiler: Reads Excel sheets exported by "CSV Parser.m" and Exports one excel sheet that contains all the data in a neat format suitable to append to a report

* GraphicsCreator: Read Excel sheets exported by "CSV Parser.m" and generates graphs to include in the reports. 

## Usage

1. Create a folder to contain all the data. 
2. Place each stations' folder in the folder that you created in step one.
3. Create a folder named 'Exported' in the folder that you created in step one. Place MonthlyTemplate.xlsx in "Exported".
4. Launch Matlab and open CSV Parser.m , ensure that the paths of rootfolder and exportfolder are correct, Also fill datafolder with the NMS folder names. Note: The code takes the last character in the folder name and cosiders it the station number(Ensure that the last number represent the station number in the folder that contains all CSV files for that station).
5. Run DataCompiler and ensure that the Exportfolder path is the same used in CSV Parser.m

