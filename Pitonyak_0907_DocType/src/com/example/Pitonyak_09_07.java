/*
 * Pitonyak_09_07.java
 *
 * Created on 2009.02.04 - 09:13:19
 *
 * Get current doc's type
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Pitonyak_09_07 {
    
    /** Creates a new instance of Pitonyak_09_07 */
    public Pitonyak_09_07() {
    }

    static private XMultiComponentFactory srvMan;
    static private XComponentContext context;
    
//    Function getDocType(Optional vDoc) As String
//        On Error GoTo Oops
//        If IsMissing(vDoc) Then vDoc = ThisComponent
//        If vDoc.SupportsService("com.sun.star.sheet.SpreadsheetDocument") Then
//            getDocType = "calc"
//        ElseIf vDoc.SupportsService("com.sun.star.text.TextDocument") Then
//            getDocType = "writer"
//        ElseIf vDoc.SupportsService("com.sun.star.drawing.DrawingDocument") Then
//            getDocType = "draw"
//        ElseIf vDoc.SupportsService(_
//            "com.sun.star.presentation.PresentationDocuments") Then
//            getDocType = "presentation"
//        ElseIf vDoc.SupportsService("com.sun.star.formula.FormulaProperties") Then
//            getDocType = "math"
//        Else
//            getDocType = "unknown"
//        End If
//    Oops:
//        If Err <> 0 Then getDocType = "unknown"
//        On Error GoTo 0 'Turn off error handling AFTER checking for an error
//    End Function

    public static String getDocType(XComponent vDoc){
        if(UnoRuntime.queryInterface(XSpreadsheetDocument.class, vDoc)!=null){
            return "calc";
        } else if (UnoRuntime.queryInterface(XTextDocument.class, vDoc)!=null){
            return "writer";
        }
        //TODO: see how tho recognize drawings, presentations and formulas
        return null;
    }
    public static String getDocType() throws Exception {
        XDesktop xDesktop =
          (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
              srvMan.createInstanceWithContext("com.sun.star.frame.Desktop",context));
        XComponent doc= xDesktop.getCurrentComponent(); // Get the current document
        return getDocType(doc);
    }

    public static void main(String[] args) {
        try {
            context = Bootstrap.bootstrap();
            srvMan = context.getServiceManager();

            XDesktop xDesktop =
              (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  srvMan.createInstanceWithContext("com.sun.star.frame.Desktop",context));

            XComponent doc= xDesktop.getCurrentComponent(); // Get the current document

            String docType= getDocType(doc);
            System.out.println("Document type: "+docType);
        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
