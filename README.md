# Auto-Text-Converter ReadME
Auto Text Converter ReadMe  (version 0.800 Beta)
March 9, 2016

The Auto Text Converter is a powerful search-and-replace tool for Microsoft Word Documents.  It utilizes regular expressions, a feature missing in standard MS Office installations, as the heart of its search-and-replace engine.  With this application one can accomplish such tasks as capturing all of the URLs in a document, converting words of British spelling into their American equivalents and vice-versa. The complexity and variety of the search and replacements one can make are only limited by one’s imagination.

Key features include:
Automated or manual controlled searching and or text replacements.
Exporting the list of searched entries to a text or EXCEL file.
The inclusion of several plugins which are used to extend the ability to perform sophisticated text replacements such as converting a date format, e.g., 3/17/2016 into March 17, 2016.
The ability of writing your own plugins and adding them to the application.
The ability of reviewing all the changes that were made during a search-and-replace session and choosing to either revert or commit that change.
One can instantly navigate to any search result by simply double-clicking the entry in the application list view.
And many more options.
How does the application work:
After starting the application, the user can choose one of several search profile XML files.  After opening the file with the application, the user can press the <process> button and the application will search paragraph by paragraph for a match.  The processing can be paused or resumed by pressing the <pause> button.  Pressing the <stop> button will stop the processing. 
What is required to run the application:
To run the application .NET 4.0 runtime must be installed.  The application was developed using Visual Studio 2015 Community edition which is a free download from the Microsoft website.  In addition MS Word 2010 or greater must be installed.  (MS Word 2007 may work, I have not tested it).

Disclaimer: (following that of MIT License)
Before running the application please read, understand, and agree to the following:
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE SaelSoft or David Saelman BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Customization
The user can create his own XML profiles for custom searches and replacements.  In this version, the user must manually create the XML file.  Existing files can be used as a template or starting point.  Below is a description of the XML format :
1. Header block – Gives general information about the XML file including version, owner and a description.
2. Search sections – Each section encapsulates a regular expression used by the document engine in performing a search. The files include:
3. 
RegEx – The regular expression. Note: one should first test the regular expression using the handy regular expression tester available from the Tools Menu.
Identifier – This field can be used for a private description of the section.
Find Color – The color the application will use for the text that will be searched for or replaced.  The list of colors which can be used can be accessed from the Tools Color Guide menu.
Action – If the action is defined as “find”, the engine will only perform a search but no replacement will be made. If the Action = “replace”, the text will be replaced.
Plugin Section – Contains fields if a plugin is used (please refer to the plugin section for more information:
PlugInName – the name of the plugin assembly
PlugInValidationFunction – the name of the function that validates the RegEx Match (optional)
PlugInFormatFunction – the name of the format function being called.
Description -- The description used to identify what the regular expression will look for.  The description appears in the RegExListBox available from the View menu.

<Searches>
  <Header>
    <Product>Auto Text Converter by SaelSoft</Product>
    <ProfileOwner>David Saelman</ProfileOwner>
    <CreationDate>March 10, 2016</CreationDate>
    <ProfileVersion>1.0</ProfileVersion>
    <ProfileDescription>Converts dates and performs other numerical formatting</ProfileDescription>
    <ProfileHistory />
  </Header>
  <Search RegEx= "(0[1-9]|1[012])[-\u2013/.](0[1-9]|[12][0-9]|3[01])[-\u2013/.]((19|20)\d\d)">
    <Identifier>URL</Identifier>
    <FindColor>Teal</FindColor>
    <Action>replace</Action>
    <Plugin Description="yes">
      <PlugInName>NumericDateValidator.dll</PlugInName>
      <PlugInValidationFunction>ValidateUSDate</PlugInValidationFunction>
      <PlugInFormatFunction>GetDateMatch_1</PlugInFormatFunction>
      <Reserved>0</Reserved>
    </Plugin>    
    <Resrved1/>
    <Resrved2/>
    <Description Text= "01/02/2008 -> January 2, 2008"/>  
  </Search>
</Searches>

Plugins 
Plugins are assemblies which are provided by this application to perform specialized replacements of text. One of the functions of plugins is to perform custom formatting that could not be done by simple regular expressions.  
A plugin can have several methods.  A XML Search section can call up to two different methods: A plugin function to perform the custom processing, and a plugin validation function to perform specialized validation of the text before formatting.  The functions must have the following prototypes:

1. Plugin function --  string MyPluginFunction(string my inputString)
2. Plugin Validation function – bool MyPluginValidationFunction(Match myRegExMatch) (optional)
If the XML section has a PluginValidation function, it will be called first. If it returns true, the plugin function will be called and the returned string will be used to replace the input.
A good example of how this works is in the SSN_Validator.dll plugin.  After a possible social security number is detected DDD DD DDDD, the validation function ValidateSSN checks to see if the numbers define a true social security number.  If they do, the GetSSNMatch function is called.  Otherwise, the number is determined to not be a valid SSN and it is skipped.

 
