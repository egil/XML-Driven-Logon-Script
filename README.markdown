The *XML Driven Logon Script* is an VBScript based logon script which behaviour (like mapping network drives, installing printers) is controlled from an xml file. This makes it easy for the none-programmer to build and manage logon scripts within his or hers organisation, since no knowledge of VBScript is needed.

Here is a rundown of the main features of the XML Driven Logon Script:

- *Separation of code and rules*. The script file (Logon.vbs) contains all the functions and objects that performs the actions defined in an xml file (logonrules.xml). Unless a new feature is needed, a change to the behaviour of the logon script is done completely in the xml file.
- *Control behaviour of the logon script by defining filters and actions in xml file(s)*. Build action sets that are selected based on one or more filter, filter by Active Directory security groups, usernames, computer names and more. Actions include ability to map network drives, add network printers, copy files and more. See the documentation for more details on available filters and actions.
- *Logging direct to the executing computers Event Log*. All information gathered by the script and all executed actions are logged to an entry in the Event Log on the computer running the script, making it easy to debug and find errors.
- *Friendly Internet Explorer window used to display messages to users*. Among other things, the logon script will display a friendly message about each action performed during execution and tell the user how long until his or hers password expire.

Take a look at the [documentation section](https://github.com/egil/XML-Driven-Logon-Script/wiki) for a more in-depth examination of the features and configuration options in the script.