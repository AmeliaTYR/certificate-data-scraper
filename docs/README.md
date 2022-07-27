# Excel Certificate Scraper User Guide

Extract data from all your downloaded pdf certificates and condense them into a easy to process and use excel sheet!

Further tools to add your certificate data to your personal website coming soon!

## Why Use this Macros?
While working on skills upgrading through online learning platforms, it becomes quickly apparent that there are few sites that can reasonably display the large library of certificates in a manner that is concise, not overwhelming to the viewer and easy to search through.

As such, extraction of the certificate data and migrating from the typical Linkedin profile built-in certificate display system to a custom certificate data display site is preferable. The idea is to condense the information while also providing links to the individual certificate proofs, so as not to subtract from the authenticity of the claims.

## Using the Macros

#### Set up

![](/images/1EnableEditing.png)

Above is the initial view of the excel sheet. If you see the protected view, you need to enable editing. This allows the macros to add data to your sheets. 

![](/images/3EnableContent.png)

After enabling editing, you may be prompted to Enable Content. This allows the macros to run.

![](/images/2LearnMoreError.png)
This may appear as you are not the author of the Macros. If you are concerned about the security of the file, the full code can be found [here](https://github.com/AmeliaTYR/certificate-data-scraper/blob/main/VBAcode.txt). To fix the error you may try to redownload the file, or follow the instructions by clicking the Learn More link.

![](/images/15Subfolders.png)

Additionally, have your certificates sorted into folder by Linkedin and Coursera. The naming of the folder does not matter, and each folder can also have subfolders as shown in the example above. The only requirements are as follows:

- All files found in the folder and its subfolder are of pdf type. 
- The folder for Coursera certificates only contains Coursera certificates.
- The folder for Linkedin certificates only contains Linkedin certificates.

Note that files of the wrong type may break the macros and prevent it from running properly.

#### Extracting File Paths Coursera

![](/images/4CouseraGetNames.png)

For the Coursera Files, first click the `Get FileNames Coursera` button. 

![](/images/5CouseraSelectFolder.png)

Next select the file containing all your Coursera certificates. 

![](/images/6CouseraNamesFound.png)

This will extract the file paths of all the pdf files in the folder and output it to the `filePathsCoursera` sheet. 

![](/images/7ExtractDataCoursera.png)

Finally, to extract the data from the pdf files found, click the `Extract PDF data Coursera` button.

![](/images/8OutputAndErrorHandling.png)

To view the output of the macros, you can view the `parsedDataCoursera` sheet. You may copy this data to any csv for further data display or processing.

#### Extracting File Paths Coursera

![](/images/9GetNamesLinkedin.png)

For the Coursera Files, first click the `Get FileNames Linkedin` button. 

![](/images/10LinkedinSelectFolder.png)

Next select the file containing all your Linkedin certificates. 

![](/images/11LinkedinNamesFound.png)

This will extract the file paths of all the pdf files in the folder and output it to the `filePathsLinkedin` sheet. 

![](/images/12ExtractDataLinkedin.png)

Finally, to extract the data from the pdf files found, click the `Extract PDF data Linkedin` button.

![](/images/13OutputAndErrorHandling.png)

To view the output of the macros, you can view the `parsedDataLinkedin` sheet. You may copy this data to any csv for further data display or processing.


#### Tips for data use
You end up with a sheet for your coursera data
and a sheet for your Linkedin data. You could display this data in a tabular form on your persnal website or Linkedin profile for ease of viewing.


## Assumptions
Assumptions about the Coursera and Linkedin certificates are made regarding the format when imported as pdf into Excel. This occasionally causes the macros to fail, and some errors may occur. 

If any outstanding macros breaking errors occur, or if you have any questions, please raise them [here](https://github.com/AmeliaTYR/certificate-data-scraper/issues/new/choose) so that I can apply the necessary fixes.

## Error handling
When the parsing goes wrong, the macros should work to do error handling and highlight rows with errors

#### Red highlights
The macros was unable to locate the required data

#### Yellow highlights
The data located is not as expected (eg the length of the title is too short)

![](/images/14ErrorYellowExample.png)

## Customization
If you have other certificate collections from other providers that have a standard format, you may request for macros updates in the repo issues [here](https://github.com/AmeliaTYR/certificate-data-scraper/issues/new/choose). 

Additionally, you may request other certificate parsing macros (eg for datacamp) in the repo issues, or ask for new fields in the certificates to be scraped (eg special notes on Linkedin certificates from secondary institutes).

## Contribute
If you would like to contribute to this macros, check out the main repo [here](https://github.com/AmeliaTYR/certificate-data-scraper) to view the code.
