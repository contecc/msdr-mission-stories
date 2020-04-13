![MSDR Cares Logo](logo.png)
# msdr-mission-stories
###PowerShell script to Dynamically Generate SharePoint Modern Pages

This script is designed to create a Modern SharePoint Article Page on a Communication Site.  It is intended to be used to fulfill an MSDR Cares project user story called "Mission Stories"  The code was generated using Patterns and Practices (PnP) PowerShell.


*Most of requirements have been met.*


####Header
- The Title and Image have been set dynamically

- [ ] AuthorByline has NOT been set.  There have been several unsuccessful attempts

####Body

There are two columns that make up the body of the article.  To create the structure, it seems a it was a little odd.  I'm not sure why I had to create a OneColumn Section after the TwoColumn section, but it worked.

```
#Add Section to the Page
Add-PnPClientSidePageSection -Page $Page -SectionTemplate TwoColumn
Add-PnPClientSidePageSection -Page $Page -SectionTemplate OneColumn -Order 2
```

######Left Column > Article Body
The left column contains the body of the article.  This accepts HTML formatted paragraphs.  

######Right Column > Story Image, Page Properties, Newsfeed
- Story Image(s):  The script puts a single image on the right column, but is anticipated that more pictures should be supported 
- [ ] Add support for multiple pictures
- Page Properties:  to help categorize the pages, the Mission Name and ID are displayed, these are vital part of the page as this will be how stories can be *collated* together
- Newsfeed:  There is a single Newsfeed WebPart added to the page that is shows all promoted articles.  It is meant to show all stories that have the same mission, is a mission was specified
- [ ] Add configuration to filter the Newsfeed Web Part to Mission Stories, OR to Mission Stories that match the ID of the page

![Screen Shot](page_screenshot.png)