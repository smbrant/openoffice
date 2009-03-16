/*
 * Pitonyak_09_06.java
 *
 * Created on 2009.02.04 - 08:08:36
 *
 * Test if a document is a text document (writer)
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Pitonyak_09_06 {
    
    /** Creates a new instance of Pitonyak_09_06 */
    public Pitonyak_09_06() {
    }

    //ch09 lis06
    //Sub DoSomethingToWriteDocument(vDoc)
    //    If NOT vDoc.supportsService("com.sun.star.text.TextDocument") Then
    //        MsgBox "A Writer document is required", 48, "Error"
    //        Exit Sub
    //    End If
    //    REM rest of the subroutine starts here
    //End Sub

    public static void main(String[] args) {
        try {
            XComponentContext context = Bootstrap.bootstrap();
            XMultiComponentFactory srvMan = context.getServiceManager();

            XDesktop xDesktop =
              (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  srvMan.createInstanceWithContext("com.sun.star.frame.Desktop",context));

            XComponent doc= xDesktop.getCurrentComponent(); // Get the current document

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
