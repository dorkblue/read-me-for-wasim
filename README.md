# read-me-for-wasim
Read me to reproduce steps for error encountered with Keyboard Simulator (KBS)

### Using pptApp.Quit() does not close powerpoint window, and subsequent actions to close slides fails

Code used:
```cs
public string Get(string id)
{
    Application pptApp = new Application();
    Presentation p = null;
    Microsoft.Office.Interop.PowerPoint.SlideShowWindows objSSWs;
    Microsoft.Office.Interop.PowerPoint.SlideShowSettings objSSS;

    if (IsPPTPresentationRunning())
    {
        // Get Running PowerPoint Application object
        pptApp = Marshal.GetActiveObject("PowerPoint.Application") as Application;
        if (pptApp != null)
        {
            p = pptApp.ActivePresentation;
        }
    }
    else
    {
        Microsoft.Office.Core.MsoTriState ofalse = Microsoft.Office.Core.MsoTriState.msoFalse;
        Microsoft.Office.Core.MsoTriState otrue = Microsoft.Office.Core.MsoTriState.msoTrue;
        pptApp.Visible = otrue;
        pptApp.Activate();
        Microsoft.Office.Interop.PowerPoint.Presentations ps = pptApp.Presentations;

        if (id.Contains("openslide="))
        {
            /*
            =================== For new clone into a new device ===================
            Please change the projectFolder string assignment below.

            IE: the project folder is the folder that houses the jarvisAPI,
            KeyBoardSimulator, and static folder.

            Copy the path of project folder and reassign below. IE:

                string projectFolder = @"< YOUR NEWLY COPIED PROJECT FOLDER PATH >";

            without the angle brackets.

            Love,
             Tom
            =======================================================================
            */
            string projectFolder = @"C:\Users\Iotuser\Documents\projects\project_jarvis";
            string slideName = id.Replace("openslide=", "");


            string pptPath = Path.GetFullPath(Path.Combine(projectFolder, "static\\powerpoints", slideName + ".pptx"));

            Debug.WriteLine(pptPath);
            p = ps.Open(@pptPath, ofalse, ofalse, otrue);
        }

        /*
            =================== Code below was originally written by Steve?  ===================
            if (id == "open1")
            {
                p = ps.Open(@"C:\Users\Iotuser\Desktop\ProjectJarvis\static\powerpoints\intro_slides.pptx", ofalse, ofalse, otrue);
            }
            else if(id == "open2")
            {
                p = ps.Open(@"C:\Users\Iotuser\Desktop\ProjectJarvis\static\powerpoints\adaptive_workforce.pptx", ofalse, ofalse, otrue);
            }
            else
            {
                p = ps.Open(@"C:\Users\Iotuser\Desktop\SmartHubSlidesToShow\Default.pptx", ofalse, ofalse, otrue);
            }
            ====================================================================================
        */

        //Run the Slide show
        objSSS = p.SlideShowSettings;
        objSSS.Run();
        objSSWs = pptApp.SlideShowWindows;
    }

    if (id == "next")
        p.SlideShowWindow.View.Next();

    if (id == "previous")
        p.SlideShowWindow.View.Previous();

    if (id == "first")
        p.SlideShowWindow.View.First();

    if (id == "last")
        p.SlideShowWindow.View.Last();

    if (id.Contains("gotoslide"))
    {
        string[] slideNuber = Regex.Split(id, @"\D+");
        p.SlideShowWindow.View.GotoSlide(Convert.ToInt32(slideNuber[1]));
    }

    if (id == "exit")
    {
        p.Close();
        pptApp.Quit();
    }


    //while (objSSWs.Count >= 1)
    //    System.Threading.Thread.Sleep(100);
    //Close the presentation without saving changes and quit PowerPoint
    //p.Close();
    //pptApp.Quit();

    #region Key Simulator Code

    //InputSimulator sim = new InputSimulator();
    //VirtualKeyCode virtualKeyCode;
    //uint KeyValue;
    //if (Enum.TryParse(id, out virtualKeyCode))
    //{
    //    KeyValue = (uint)virtualKeyCode;
    //}
    ////VirtualKeyCode VRKey = (VirtualKeyCode)System.Enum.Parse(typeof(VirtualKeyCode), id);

    //sim.Keyboard.KeyPress(VirtualKeyCode.VOLUME_MUTE);//(virtualKeyCode);//

    ////Open Notepad and write on it
    //Process notepad = new Process();
    //notepad.StartInfo.FileName = "notepad.exe";
    //notepad.Start();
    //notepad.WaitForInputIdle();
    //IntPtr notepadHandle;
    //notepadHandle = notepad.MainWindowHandle;
    ////write in the notepad
    //sim.Keyboard.TextEntry(id);

    //// CTRL + C (effectively a copy command in many situations)
    //sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_C);

    //// CTRL + V (effectively a paste command in many situations)
    //sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
    ////notepad.CloseMainWindow();

    //Process[] explorer = Process.GetProcessesByName("iexplore");
    //foreach (Process ie in explorer)
    //{
    //    ie.WaitForInputIdle();
    //    sim.Keyboard.KeyPress(VirtualKeyCode.F5);
    //}
    ////IntPtr ptr = explorer.MainWindowHandle;
    ////explorer.WaitForInputIdle();

    #endregion

    return "value";
}
```

Steps to reproduce:
> 1. Trigger slide open - Success
> 2. Trigger slide close - Success (but ppt window still opened)
> 3. Trigger slide open - Success
> 4. Trigger slide close - Failed (slideshow does not close, ppt window does not close, error below returned to Nodejs)

Error returned to Node js server
```javascript
{ Message: 'An error has occurred.',
  ExceptionMessage: 'Operation unavailable (Exception from HRESULT: 0x800401E3 (MK_E_UNAVAILABLE))',
  ExceptionType: 'System.Runtime.InteropServices.COMException',
  StackTrace: '   at System.Runtime.InteropServices.Marshal.GetActiveObject(Guid& rclsid, IntPtr reserved, Object& ppunk)\r\n   at System.Runtime.InteropServices.Marshal.GetActiveObject(String progID)\r\n   at KeyBoardSimulator.Controllers.ValuesController.Get(String id) in C:\\Users\\Iotuser\\Documents\\projects\\project_jarvis\\KeyBoardSimulator\\KeyBoardSimulator\\KeyBoardSimulator\\Controllers\\ValuesController.cs:line 40\r\n   at lambda_method(Closure , Object , Object[] )\r\n   at System.Web.Http.Controllers.ReflectedHttpActionDescriptor.ActionExecutor.<>c__DisplayClass6_1.<GetExecutor>b__3(Object instance, Object[] methodParameters)\r\n   at System.Web.Http.Controllers.ReflectedHttpActionDescriptor.ActionExecutor.Execute(Object instance, Object[] arguments)\r\n   at System.Web.Http.Controllers.ReflectedHttpActionDescriptor.ExecuteAsync(HttpControllerContext controllerContext, IDictionary`2 arguments, CancellationToken cancellationToken)\r\n--- End of stack trace from previous location where exception was thrown ---\r\n   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task)\r\n   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task)\r\n   at System.Web.Http.Controllers.ApiControllerActionInvoker.<InvokeActionAsyncCore>d__1.MoveNext()\r\n--- End of stack trace from previous location where exception was thrown ---\r\n   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task)\r\n   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task)\r\n   at System.Web.Http.Controllers.ActionFilterResult.<ExecuteAsync>d__5.MoveNext()\r\n--- End of stack trace from previous location where exception was thrown ---\r\n   at System.Runtime.CompilerServices.TaskAwaiter.ThrowForNonSuccess(Task task)\r\n   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task)\r\n   at System.Web.Http.Dispatcher.HttpControllerDispatcher.<SendAsync>d__15.MoveNext()' }
```

### Taskkill Video.UI crashes KBS on Visual Studio
Code used:
```javascript
shell.exec("taskkill /F /IM Video.UI.exe")
```

Steps to reproduce:
> 1. Open video thru api endpoint `localhost:3000/api/videos/open/:name`
>
>       replace :name with any of the following 
>       - british_gas
>       - tata_steel
>       - safe_city
> 2. Close video thru api endpoint `localhost:3000/api/videos/close`

I'm unable to replicate the bug that we faced that day. For some reason right now closing the video with taskkill does not crash the KBS.

But if it is of any help, the error message that I've saw previously was produced by `iisexpress.exe` in the output window, and has something to do with `Cannot find or open the PDB file`.
