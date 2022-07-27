# Excel Certificate Scraper User Guide

Extract data from all your downloaded pdf certificates and condense them into a easy to process and use excel sheet!

Further tools to add your certificate data to your personal website coming soon!

## Using the Macros

#### Set up

![](/images/1EnableEditing.png)

- This is the initial view

![](/images/3EnableContent.png)

![](/images/2LearnMoreError.png)
- If you see the button select enable macros

- Also have your certificates sorted into folder by Linkedin and Coursera


#### Extracting File Paths
![](/images/4CouseraGetNames.png)
![](/images/5CouseraSelectFolder.png)
![](/images/6CouseraNamesFound.png)
![](/images/7ExtractDataCoursera.png)
![](/images/8OutputAndErrorHandling.png)
![](/images/9GetNamesLinkedin.png)
![](/images/10LinkedinSelectFolder.png)
![](/images/11LinkedinNamesFound.png)
![](/images/12ExtractDataLinkedin.png)
![](/images/13OutputAndErrorHandling.png)

To extract all certificate file names and file paths, use the first button

Similarly for Linkedin

#### Parsing data
Press the next button

#### Tips for data use
You end up with a sheet for your coursera data
and a sheet for your Linkedin data


## Assumptions
- Assumptions about the coursera cert format
- Assumptions about the 

## Error handling
When the parsing goes wrong, the macros should work to do error handling and highlight rows with errors

#### Red highlights
The macros was unable to locate the required data

#### Yellow highlights
The data located is not as expected (eg the length of the title is too short)

## Customization