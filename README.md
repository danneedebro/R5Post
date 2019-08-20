# R5Post
Spreadsheet based tool to batch run RELAP5 calculations and subsequent post-processing

![Layout](Images/R5Post-description.png)


## Idea
By using the timestamps of the file created during the calculation and post-processing chain it's easy to see what files/cases are up to date. If they are out of date it is easy to control the work flow from one place.

## Structure
Place R5Post.xlsm in a parent folder to your cases. Then write the relative path to the RELAP5 input file (Case1\Case1.i) and all files that are to be created is lit up automatically.
