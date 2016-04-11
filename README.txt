_______________________________________________________________
OVERVIEW

The ElancoPimsDdsParser console application parses a Word document and
generates Excel data.

The ELANCO customer uses Word to describe the details of the PIMS batch
report. ELANCO also transcribes these details into Excel sheets that
are then used by Aspen Batch.21 tools to create Batch reports.
(The ELANCO facility is in Augusta, GA.)

This program can assist the manual process of transcribing the PIMS 
report details from the Word DDS (Detail Design Specification) to 
Excel.

This program uses the the MS OpenXML SDK (v 2.5) to parse the Word doc
data in XML format. (I believe the SDK author is a "Chris White").

_______________________________________________________________
USAGE

This program is a command line utitily that expects the input file
name of the Word document. This program outputs an Excel document 
named with the same basename as the input file.

_______________________________________________________________
CAVEATS

To achieve the expected output the Word document must:
  1) have all track-changes "accepted"
  2) NOT have any tables with their data split among multiple
     tables.
	 
For #2 above, each "Batch Area Characteristic Processing" table 
contains the data for a specific function. Original ELANCO docs would
sometimes have the data for one function split among multiple tables
for aesthetic reasons. However for the parsing logic to work correctly
(at least with this version), that data must be contained within a 
single table.

