#Configuration ###################################################################################################

$SiteURL = "https://microsoft.sharepoint.com/teams/disaster"  #url of the aka.ms/disasterinfo site  (shouldn't change, but just in case)

$associatedMission = "Aus-net Australian Brush Fires"  #is the name of the ACTUAL Mission from our Database, this may change  

$associatedMissionId = "FY20-14" #The Mission NUMBER that shouldn't change over time

$releasable = $true    #is $true or $false depending on the user > the default in SharePoint column is NO, I'm setting to use to PROVE it is being set

$PageName = "MissionStory - " + $associatedMission #Should Either be same as the mission name or perhaps a Sequential Value (ie. Mission Story 001, 002, 003, etc00)

$PageTitle = "My Red Cross Adventure in Australia"  #This is up to the Volunteer to Title their own story

#created / Modified BY (Author)
$authorUPN = "lcurtis@microsoft.com"

#Page Body 
#DOES accept HTML for formatting paragraphs
$PageBody = "<h4>submitted by: " + $authorUPN + "</h4>" + "<p>lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci. Aenean nec lorem.
In porttitor. Donec laoreet nonummy augue. Suspendisse dui purus, scelerisque at, vulputate vitae, pretium mattis, nunc. Mauris eget neque at sem venenatis eleifend. Ut nonummy. Fusce aliquet pede non pede. Suspendisse dapibus lorem pellentesque magna. Integer nulla. Donec blandit feugiat ligula. Donec hendrerit, felis et imperdiet euismod, purus ipsum pretium metus, in lacinia nulla nisl eget sapien.
Donec ut est in lectus consequat consequat. Etiam eget dui.</p><br/> <p>Aliquam erat volutpat. Sed at lorem in nunc porta tristique. Proin nec augue. Quisque aliquam tempor magna. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Nunc ac magna. Maecenas odio dolor, vulputate vel, auctor ac, accumsan id, felis. Pellentesque cursus sagittis felis.
Pellentesque porttitor, velit lacinia egestas auctor, diam eros tempus arcu, nec vulputate augue magna vel risus.</p> <br/> <p>Cras non magna vel ante adipiscing rhoncus. Vivamus a mi. Morbi neque. Aliquam erat volutpat. Integer ultrices lobortis eros. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin semper, ante vitae sollicitudin posuere, metus quam iaculis nibh, vitae scelerisque nunc massa eget pede. Sed velit urna, interdum vel, ultricies vel, faucibus at, quam. Donec elit est, consectetuer eget, consequat quis, tempus quis, wisi.
In in nunc. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos hymenaeos. Donec ullamcorper fringilla eros. Fusce in sapien eu purus dapibus commodo. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Cras faucibus condimentum odio. Sed ac ligula. Aliquam at eros. Etiam at ligula et tellus ullamcorper ultrices. In fermentum, lorem non cursus porttitor, diam urna accumsan lacus, sed interdum wisi nibh nec nisl.
Ut tincidunt volutpat urna. Mauris eleifend nulla eget mauris. Sed cursus quam id felis. Curabitur posuere quam vel nibh. Cras dapibus dapibus nisl. Vestibulum quis dolor a felis congue vehicula. Maecenas pede purus, tristique ac, tempus eget, egestas quis, mauris. Curabitur non eros. Nullam hendrerit bibendum justo. Fusce iaculis, est quis lacinia pretium, pede metus molestie lacus, at gravida wisi ante at libero.
</p>"  #This Body text will get replace with the volunteer's story  

#Page Header Background Image
#$PageHeaderImage = "/teams/Disaster/SiteAssets/SitePages/AustralianWildfires/fires2.jpg"  #Server Relative URL 
$PageHeaderImage = "/teams/Disaster/SiteAssets/SitePages/Disaster/fire.jpg"  #Server Relative URL 

$articlePicture = "https://microsoft.sharepoint.com/teams/Disaster/SiteAssets/SitePages/Disaster/AustralianFires.jpg"  #FQDN plus path
##################################################################################################################


#Connect to PpP Online
Connect-PnPOnline -Url $SiteURL -UseWebLogin # -Credentials (Get-Credential)
  
#Create new page with the Layout of Article
$Page = Add-PnPClientSidePage -Name $PageName -LayoutType Article

Set-PnPClientSidePage -Identity $Page -Title $PageTitle -CommentsEnabled:$true -HeaderType Custom -ServerRelativeImageUrl $PageHeaderImage -LayoutType Article


#AUTHOR
#TODO > Set the Page Author to the Person who submitted the Story
#Using Code to Add an Author / Byline to a Modern Page in SharePoint Online
#https://threewill.com/using-code-to-add-an-author-byline-to-a-modern-page-in-sharepoint-online/
#https://techcommunity.microsoft.com/t5/sharepoint/resolved-how-to-change-modern-page-author-from-the-quot-created/m-p/220432

#Author

#Associated Mission
$Page.PageListItem.Properties["AssociatedMission"] = $associatedMission
$Page.PageListItem.Update()

#Associated Mission
$Page.PageListItem.Properties["AssociatedMissionId"] = $associatedMissionId
$Page.PageListItem.Update()

#Associated Mission
$Page.PageListItem.Properties["Releasable"] = $true
$Page.PageListItem.Update()


#Add Section to the Page
Add-PnPClientSidePageSection -Page $Page -SectionTemplate TwoColumn
Add-PnPClientSidePageSection -Page $Page -SectionTemplate OneColumn -Order 2

#Not sure why this combination of sections works, but it does:  Leveraging. https://hangconsult.com/2017/11/05/creating-a-new-client-side-page-with-pnp-powershell/

#Column 1
#Add Story (Body) to Page
Add-PnPClientSideText -Page $Page -Text $PageBody -Section 2 -Column 1

#Column 2 #####################################################################
#has the MetaData, Images and Links to other 'related' stories

#Page Properties
Add-PnPClientSideWebPart -Page $Page -DefaultWebPartType PageFields -WebPartProperties @{selectedFieldIds = "33a46e28-aeea-4f81-ae7e-a60a34e606d4","43687d46-835b-427e-8ba0-0ca74c4acd2f"} -Section 2 -Column 2

#Image(s)
#won't display unless the height and width of the image are also set.  
#It *may* make some sense to pre-configure 3-5 pre-configured HTML DIVs with some CSS and set the background of the DIVs so that the Images always maintain the appropriate aspect ratio
Add-PnPClientSideWebPart -Page $Page -DefaultWebPartType Image -WebPartProperties @{imageSource=$articlePicture;imgWidth=300;imgHeight=200 } -Section 2 -Column 2 

#Add News web part to the section  (Need to add the Filter > Property = Associated Mission Id equals Whatever the ID of the Mission being created.  IOW, Stories with the same missionId)
Add-PnPClientSideWebPart -Page $Page -DefaultWebPartType NewsFeed -Section 2 -Column 2
 
#end Column 2 #################################################################

#Save/Publish the Page
$Page.Save()
$Page.Publish()

#IDK WHY I had to set the images AFTER the page was created, but IT WORKS!!!! Woot Woot!!!
#Set the Page Header Image
$Page.SetCustomPageHeader($PageHeaderImage)  #from https://www.sharepointdiary.com/2019/03/sharepoint-online-change-banner-image-in-modern-site-pages.html
$Page.Save()
$Page.Publish()


#Updating the Created By to the ACTUAL Author (Whomever submitted the story via the form)
$pageNamePlusExt = $PageName + ".aspx"
$pageTemp = Get-PnPClientSidePage $pageNamePlusExt 
Set-PnPListItem -List "Site Pages" -Identity $pageTemp.PageListItem.Id -Values @{"Editor"=$authorUPN;"Author"=$authorUPN;"_AuthorByline"=$authorUPN}
Set-PnPClientSidePage $pageNamePlusExt -Publish

<#
After adding this section to update the author, the script does throw an error, but still appears to work 
format-default : The collection has not been initialized. It has not been requested or the request has not been executed. It may need to be explicitly requested.
    + CategoryInfo          : NotSpecified: (:) [format-default], CollectionNotInitializedException
    + FullyQualifiedErrorId : Microsoft.SharePoint.Client.CollectionNotInitializedException,Microsoft.PowerShell.Commands.FormatDefaultCommand
#>