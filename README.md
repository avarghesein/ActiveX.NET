# ActiveX.NET
A true Out-Of-Proc (EXE) COM server implemented in C# Windows Application based on [MEF Plugin](https://docs.microsoft.com/en-us/dotnet/framework/mef/attributed-programming-model-overview-mef) Architecture. Develop COM objects in C# Libraries, which automatically get hosted in EXE COM Server, with minimal configuration.

Originally inspired from [CSExeCOMServer](https://code.msdn.microsoft.com/windowsapps/CSExeCOMServer-3b1c1054) published in https://code.msdn.microsoft.com . The idea is to provide a more advanced implementation, with the below features:

1. Replace Native Windows Message Loop, with .NET Windows Message Loop (Application.Run)

       This will enable leveraging .NET Win Forms/UserControls inside COM Objects created within .NET 

2. Decouple Out-Of-Process EXE Server and make it as a shared EXE COM Server

       This will enable loading user created COM Visible Libraries dynamically

3. Help developers build COM Objects easily, targeting for EXE COM Server

       Let developer build COM objects without worrying about reference counting, and registering for Out-Of-Proc use. Enable the COM Object for Out-Of-Proc use by simply applying a few attributes, base classes.


![alt Architecture](https://github.com/avarghesein/ActiveX.NET/blob/master/ActiveX.NET.Architecture.jpg)

HOW TO RUN THE PROJECT?

       a. Run Visual Studio As Administrator
       b. Make "ActiveX.NET.Server" as startup project
       c. Ensure CommandLine Arguments value to "/regserver" in Debug Tab
       d. Hit F5 To Run the selected project, and wait until it finishes execution
       e. Open the excel (Enable Macros While Opening) and open its VBA Editor, and run the sample VBA Macro.
       f. You should be viewing Hello World Message from VBA
       
   ![alt RunningSource](https://github.com/avarghesein/ActiveX.NET/blob/master/RunningSource.jpg)
       


HOW TO DEVELOP YOUR OWN COM-PLUGIN DLL(S)?

a.	Create your COM Visible Plug-In DLL(s)

     Refer Sample: “ActiveX.NET.Plugin” project

    For all COM Objects that should be used as Out Of Proc, ensure the below:

      i.	Apply “Guid” attribute
  
      ii.	Apply “ActiveXServer” attribute
  
      iii.	Inherit from “ActiveXServerBase”

   ![alt Sample](https://github.com/avarghesein/ActiveX.NET/blob/master/COMObjectForOutOfProcSample.JPG)
  

b.	Copy the Plugin-In DLL(s), to the configured Plug-In Location

     Location should be configured in the ActiveX.NET.Server config/appsetting: “ActiveXServerPlugins:Location”
  

c.	Register Plugin-In DLL(s) for COM Use

      This can be done by registering the ActiveX.NET.Server using the below command line. The server will automatically load all available Plug-In DLL’s available in Plug-In Location and register them for Out Of Proc Use

      Sample: C:> ActiveX.NET.Server.exe /regserver
   

d.	Now Use any COM Clients to instantiate the COM Object

      Refer “OutOfProcTestUsingVBA.xlsm”
   ![alt Sample](https://github.com/avarghesein/ActiveX.NET/blob/master/TestOutOfProcComObject.JPG)
   

e.	To Unregister Plug-In DLL(S)

      Sample: C:> ActiveX.NET.Server.exe /unregserver


# What's Next?
Try create a ComVisible Windows User Control, that could be used from ComClients like VB.
You may refer one sample [here](https://haseebakhtar.wordpress.com/2011/05/31/creating-an-activex-control-in-net-using-c/)
