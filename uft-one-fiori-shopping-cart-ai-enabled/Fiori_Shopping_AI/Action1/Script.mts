'=================================================================================================================================================================================
'	Note that the demo application has the text "Available" on both the left frame, list of products, and in the middle, advertisement frame.
'		This is why we're using micFromLeft rather than miFromTop, because some categories will have the 2nd available product below the advertisement frame products
'=================================================================================================================================================================================
Dim Category, CategoryListHeader, rc, oShell											'Initialize the variables

Set oShell = CreateObject ("WSCript.shell")
oShell.run "powershell -command ""Start-Service mediaserver"""
Set oShell = Nothing

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3													'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")													'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

'================================================================================================
'This code will make it so that the script will be able to be run in both 15.0.1 and 15.0.2+ environment
If isempty(micAnyText) and not isempty(micNoText) Then
    micAnyText = micNoText    
End If
'================================================================================================

Category = DataTable.GlobalSheet.GetParameter("Categories")						'Set the value for the Category that will be clicked on
CategoryListHeader = "< " & Category											'Set the value for the Category header in the list of products
AIUtil.FindTextBlock(Category).Click											'Click the value in the datasheet in the category menu, originally created with the Laptops category
'=================================================================================================================================================================================
'	Example of an AI sync point
'=================================================================================================================================================================================
rc = AIUtil.FindTextBlock("EUR", micFromTop, 5).Exist(20)							'Wait for the page to load to show th text EUR 5 times on the screen (so the left pane list of products has loaded)
AIUtil.FindTextBlock("Available", micFromLeft, 1).Click							'Click on the first available product 
AIUtil.FindTextBlock("Available", micFromLeft, 2).Click							'Click on the second available product
AIUtil("button", "Add to Cart").Click											'Click on the Add to Cart button.
AIUtil("shopping_cart").Click													'Click the shopping cart icon
AIUtil.FindTextBlock(CategoryListHeader).Click 50, 1							'Click on the text of the category header to allow the application to catch up, could replace with a .Highlight
AIUtil("pencil").Click															'Click the edit icon, shaped like a pencil
AIUtil.FindTextBlock(CategoryListHeader).Click 50, 1							'Click on the text of the category header to allow the application to catch up, could replace with a .Highlight
Browser("Browser").Maximize														'Maximize the browser or the objects won't be visible
AIUtil("close").Click															'Click the delete button for the first item in the cart, script assumes there is only one item in the cart
AIUtil("button", "Delete").Click												'Click the Delete button in the pop-up frame
AIUtil.FindTextBlock("Save Changes").Click										'Click the Save Changes in the cart slide out frame
AIUtil("left_triangle", micAnyText, micFromTop, 1).Click							'Click the arrow next to the category header to move back to the main categories page
AIUtil("shopping_cart", micAnyText, micFromTop, 1).Click													'Click the cart icon to collapse the shopping cart slide out frame.
'AIUtil("button", "").Click														'Click the cart icon to collapse the shopping cart slide out frame.
Browser("Browser").Close														'Close the browser window
