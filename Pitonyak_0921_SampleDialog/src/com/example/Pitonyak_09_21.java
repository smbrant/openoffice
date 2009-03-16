/*
 * Pitonyak_09_21.java
 *
 * Created on 2009.02.04 - 14:41:06
 *
 */

package com.example;

import com.sun.star.awt.XDialogProvider;
import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.document.XEmbeddedScripts;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Pitonyak_09_21 {
    
    /** Creates a new instance of Pitonyak_09_21 */
    public Pitonyak_09_21() {
    }

    //Private MySampleDialog As Object
    //Sub DisplayObjectInformation(Optional vOptionalObj)
    //    Dim vControl 'I access the text control in the dialog
    //    Dim s$ 'Temporary string variable
    //    Dim vObj 'Object about which to display information
    //    REM If ths object is not provided, use the current document
    //    If IsMissing(vOptionalObj) Then
    //        vObj = ThisComponent
    //    Else
    //        vObj = vOptionalObj
    //    End If
    //    REM Create the dialog and set the title
    //    MySampleDialog = CreateUnoDialog(DialogLibraries.OOMECH09.MyFirstDialog)
    //    MySampleDialog.setTitle("Variable Type " & TypeName(vObj))
    //    REM Get the text field from the dialog
    //    REM I added this text manually
    //    vControl = MySampleDialog.getControl("TextField1")
    //    If InStr(TypeName(vObj), "Object") < 1 Then
    //        REM If this is NOT an object, simply display simple information
    //        vControl.setText(Dlg_GetObjTypeInfo(vObj))
    //    ElseIf NOT HasUnoInterfaces(vObj, "com.sun.star.uno.XInterface") Then
    //        REM It is an object but it is not a UNO object
    //        REM I cannot call HasUnoInterfaces if it is not an object
    //        vControl.setText(Dlg_GetObjTypeInfo(vObj))
    //    Else
    //        REM This is a UNO object so attempt to access the "dbg_" properties
    //        MySampleDialog.setTitle("Variable Type " & vObj.getImplementationName())
    //        s = "*************** Methods **************" & CHR$(10) &_
    //        Dlg_DisplayDbgInfoStr(vObj.dbg_methods, ";") & CHR$(10) &_
    //        "*************** Properties **************" & CHR$(10) &_
    //        Dlg_DisplayDbgInfoStr(vObj.dbg_properties, ";") & CHR$(10) &_
    //        "*************** Services **************" & CHR$(10) &_
    //        Dlg_DisplayDbgInfoStr(vObj.dbg_supportedInterfaces, CHR$(10))
    //        vControl.setText(s)
    //        End If
    //    REM tell the dialog to start itself
    //    MySampleDialog.execute()
    //End Sub
    public static void main(String[] args) {
        try {
            XComponentContext context = Bootstrap.bootstrap();
            XMultiComponentFactory srvMan = context.getServiceManager();

            XDesktop xDesktop =
              (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  srvMan.createInstanceWithContext("com.sun.star.frame.Desktop",context));

            XComponent doc= xDesktop.getCurrentComponent(); // Get the current document

            XDialogProvider dialogProvider= (XDialogProvider) UnoRuntime.queryInterface(XDialogProvider.class,
                    srvMan.createInstanceWithContext("com.sun.star.awt.DialogProvider", context));

//            XDialog dialog= dialogProvider.createDialog("Dialog1");
//            dialog.execute();

            XEmbeddedScripts embeddedScripts= (XEmbeddedScripts) UnoRuntime.queryInterface(XEmbeddedScripts.class,
                    srvMan.createInstanceWithContext("com.sun.star.document.EmbeddedScripts", context));

            //TODO: have an error here!
            String[] dialogs=embeddedScripts.getDialogLibraries().getElementNames();

        for (String dial : dialogs)
            System.out.println(dial);

            XTextDocument text= (XTextDocument)UnoRuntime.queryInterface(XTextDocument.class, doc);

            if (text==null){
                System.out.println("A Writer document is required.");
            }else{
                System.out.println("The current document is a Writer document.");
            }
        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
}
