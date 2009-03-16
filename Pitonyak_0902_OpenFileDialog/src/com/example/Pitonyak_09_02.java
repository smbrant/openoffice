/*
 * Pitonyak_09_02.java
 *
 * Created on 2009.02.02 - 15:21:37
 *
 * File dialog
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.ucb.XSimpleFileAccess;
import com.sun.star.ui.dialogs.XFilePicker;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Pitonyak_09_02 {
    
    /** Creates a new instance of Pitonyak_09_02 */
    public Pitonyak_09_02() {
    }
    
    //ch09 lis02
    //Function ChooseAFileName() As String
    //    Dim vFileAccess 'SimpleFileAccess service instance
    //    Dim iAccept as Integer 'Response to the FilePicker
    //    Dim sInitPath as String 'Hold the initial path
    //    'Note: The following services MUST be called in the following order
    //    'or Basic will not remove the FileDialog Service
    //    vFileDialog = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
    //    vFileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    //    'Set the initial path here!
    //    sInitPath = ConvertToUrl(CurDir)
    //    If vFileAccess.Exists(sInitPath) Then
    //      vFileDialog.SetDisplayDirectory(sInitPath)
    //    End If
    //    iAccept = vFileDialog.Execute() 'Run the file chooser dialog
    //    If iAccept = 1 Then 'What was the return value?
    //    GetAFileName = vFileDialog.Files(0) 'Set file name if it was not canceled
    //    End If
    //    vFileDialog.Dispose() 'Dispose of the dialog
    //End Function
    public static void main(String[] args) throws Exception {

        XComponentContext context = Bootstrap.bootstrap();
        XMultiComponentFactory srvMan = context.getServiceManager();

        XFilePicker vFileDialog=
                (XFilePicker) UnoRuntime.queryInterface(XFilePicker.class,
                srvMan.createInstanceWithContext("com.sun.star.ui.dialogs.FilePicker", context));

        XSimpleFileAccess vFileAccess=
                (XSimpleFileAccess) UnoRuntime.queryInterface(XSimpleFileAccess.class,
                srvMan.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", context));

        String sInitPath = "c:\\temp";
        if (vFileAccess.exists(sInitPath)){
            vFileDialog.setDisplayDirectory(sInitPath);
        }
        short iAccept= vFileDialog.execute();
        if (iAccept == 1){
            String[] files= vFileDialog.getFiles();
            System.out.println(files[0]);
        }

        System.exit( 0 );
    }
    
}
