# Excel Certificate scraper
An excel macros automation created to extract information from Coursera and Linkedin PDF 
files and consolidate them in an excel sheet format.

## Motivation
While working on skills upgrading through online learning platforms, it becomes quickly apparent
that there are few sites that can reasonably display the large library of certificates in a manner 
that is concise, not overwhelming to the viewer and easy to search through.

As such, extraction of the certificate data and migrating from the typical Linkedin profile built-in
certificate display system to a custom certificate data display site is preferable. The idea is to 
condense the information while also providing links to the individual certificate proofs, so as not 
to subtract from the authenticity of the claims.

## Future improvments

### Reduction of hard coded parsing functions
Currently, some of the macros code is hard coded to the format of certificates as of Jul 2022. 
Updates to the parsing functions would be needed if either of the certificate formats are 
updated by their providers. 

### Additional user input
Additional customaization of which data to extract from these sheets would make the code more 
reusable in the long run.

### More comprehensive error handling
Some certificates contain extra data such as external certifying organizations or credits. Additional 
error handling is required for when the search parameters are not found within the document.

### Parsing of other Certificate types
Other certificate formats could potentially be added. Sample certificate data is required.

## Longest common substring
Use of longest common substring to sort the certificates by category automatically based on folder

## Further search for patterns 
For more dynamic parsing of certs where there are variations in the import

## Review of assumptions
Proper review and documentation of the hard coded segments to better identify

## Potential use of OpenCV for text extraction
To double check, may be possible to also run an opencv analysis of the file to identify the text and 
cross refer to reduce the error rates

## Skipping of empty cells
If information searched for is not found, may want to skip empty cells
For the title if the title is shorter than expected, try checking if the cell to the right is empty or contains 
the rest of the text like the bugged example

## Update to include parsing of Top Skills Covered
Additional field discovered in Linkedin Pathway search

## Companion web display template
I am currently working on a website template to display the newly extracted data.

## User guide
> User guide coming soon!

## Open source
This project is formally open for public use and contributions. Additionally, you may open issues if the code
does not run as expected (ie report bugs)

## Credits
The code for looping through a folder and subfolders is adapted from [ExtendOffice](https://www.extendoffice.com/documents/excel/2994-excel-list-all-files-in-folder-and-subfolders.html) 
and modified accordingly:

- In addition to listing the file names, the full path is listed as well
- Added parameter to determine which sheet the function would output the data to

## Conclusion
This project has given me hands on experience with writing VBA macros, and I learned a lot about
the extent to which Macros are able to automate repetitive tasks such as PDF scraping. I hope to find 
more tasks to automate with VBA.

