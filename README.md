# spfadsync
<p><strong>Project Description</strong></p>
<p>The objective of this project is to provide a simple and effective way to synchronize the properties of a SharePoint user with the properties of their domain account.</p>
<p>Indeed, in foundation version of SharePoint, information in the UserInformation list only syncs with AD when the user is first added or logs in the first time.&nbsp;No synchronization properties is provided in this release, it is only available with the
 service application &quot;User Profile&quot; available in the paid version (SharePoint Server).</p>
<p>However, there is a native ability to synchronize accounts with the cmdlet &quot;Set-SPUser&quot; and the parameter &quot;SyncFromAD&quot;. However, it only updates the name (Name / Display name) and email address (E-mail / Work E-mail).</p>
<p>To go further, it is necessary to directly update the list &quot;User Information List&quot; with the attributes of the accounts. The account attributes are easily retrievable via cmdlets &quot;Get-ADUser&quot; provided in the &quot;Active Directory for Windows PowerShell module&quot;
 feature available with Windows 2008 R2 or higher.</p>
<p><strong>Audience</strong></p>
<p>The script was written for sharePoint administrators who want to synchronize SharePoint User Profile of SharePoint Foundation farm with Active Directory information.</p>
<p><strong>Features</strong></p>
<p>I designed a script that allows you to:</p>
<ul>
<li>Treat all users of all site collections in each of the web applications (Basic and claim account)
</li><li>Check the availability of the domain of the user </li><li>Possiblity to add forest name and credential&nbsp;if the account is from a different forest than farm
</li><li>Sync user with native cmdlet (Set-SPUser with SyncFromAD parameter) </li><li>Check if the user is in the domain and if it has been modified or recreated </li><li>Update the user in SharePoint if it has been modified or recreated (Move-SPUser with IgnoreSID parameter)
</li><li>Synchronize Job title, Department, IPPhone, Mobile Phone and Title attributes (Only with&nbsp;Windows 2008 R2 or higher)
</li><li>Check if the attributes have been changed </li><li>Delete accounts that are not found or in an unreachable domain (Only if the number of deleted accounts is less than 30% of accounts synchronized)
</li><li>Logging all actions performed </li><li>Send a detailed report by email </li></ul>
<p>This script has been tested successfully with :</p>
<ul>
<li>Windows 2008, Powershell V2 and a SharePoint Foundation 2010 farm with several hundred&nbsp;users froma same domain as SharePoint farm
</li><li>Windows 2008 R2, Powershell V2 and a SharePoint Foundation 2010 farm with several thousand users from several domain in the same forest&nbsp;as SharePoint farm
</li><li>Windows 2012 R2, Powershell V4 and a SharePoint Foundation 2013 farm with several hundred&nbsp;users from same domain as SharePoint farm and several domain in the another forest (One-way trust)
</li></ul>
<p>Of course, this script is not perfect and it could be better written, do not hesitate to send me your feedback.</p>
<p><strong>Prerequisites</strong></p>
<p>The script must be run on the SharePoint Server (2010/2013).</p>
<p>The script is fully functional by installing&nbsp;&quot;Active Directory for Windows PowerShell module&quot; feature available as part of the Remote Server Administration Tools (RSAT) feature on a Windows&nbsp;Server&nbsp;2008&nbsp;R2 server or higher.</p>
<p>Your Active Directory accounts must be up to date&nbsp;to not replace the information entered by users with information that is outdated.</p>
<p>Edit Variable configuration on the top of the script before running it.</p>
<p>If you want to test synchronization on a single web application or site collection, you can change the 1278 line of the script by replacing
<code>$sites = Get-SPSite -Limit ALL</code>&nbsp;with <code>$sites = Get-SPSite http://yoursiteurl</code></p>
<p><strong>References</strong></p>
<p><a title="Updating SharePoint 2010 User Information" href="http://blog.falchionconsulting.com/index.php/2011/12/updating-sharepoint-2010-user-information/">Updating SharePoint 2010 User Information</a></p>
<p><a title="Sharepoint Foundation 2010 MAJ avec AD" href="https://social.technet.microsoft.com/Forums/fr-FR/f464d4f6-4c30-48cb-871c-e78d1445f940/sharepoint-foundation-2010-maj-avec-ad?forum=sharepoint2010tnfr">Sharepoint Foundation 2010 MAJ avec AD</a></p>
<p><a title="Account SID’s" href="http://blogs.technet.com/b/marios_mo_betta_blog/archive/2012/07/05/account-sid-s.aspx">Account SID&rsquo;s</a></p>
<p><a title="Set-SPUser" href="https://technet.microsoft.com/en-us/library/ff607827(v=office.14).aspx">Set-SPUser</a></p>
<p><a title="Move-SPUser" href="https://technet.microsoft.com/en-us/library/ff607729(v=office.14).aspx">Move-SPUser</a></p>
<p><a title="Get-ADUser" href="https://technet.microsoft.com/en-us/library/ee617241.aspx">Get-ADUser</a></p>
<p><a title="Active Directory: Get-ADUser Default and Extended Properties" href="http://social.technet.microsoft.com/wiki/contents/articles/12037.active-directory-get-aduser-default-and-extended-properties.aspx">Active Directory: Get-ADUser Default and Extended
 Properties</a></p>
<p><strong>Known errors</strong></p>
<p>Move-SPUser Failed with error &quot;The site with the id &quot;GUID&quot; could not be found :</p>
<p><a href="http://wscheema.com/blog/Lists/Posts/Post.aspx?ID=31">http://wscheema.com/blog/Lists/Posts/Post.aspx?ID=31</a></p>
<p>&nbsp;</p>
