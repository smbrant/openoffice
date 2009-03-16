/*
 * Macro11_paragraph_styles.java
 *
 * Created on 2009.02.09 - 15:02:22
 *
 * Lists paragraph styles in document
 *
 */

package com.example;

import com.sun.star.beans.XPropertySet;
import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Macro11_paragraph_styles {
    
    /** Creates a new instance of Macro11_paragraph_styles */
    public Macro11_paragraph_styles() {
    }

// Original from http://codesnippets.services.openoffice.org/Writer/Writer.ParagraphTraversal.snip

//Sub PrintAllParagraphStyles
//  Dim s As String
//  Dim vCurCursor As Variant
//  Dim vText As Variant
//  Dim sCurStyle As String
//
//  vText = ThisComponent.Text
//  vCurCursor = vText.CreateTextCursor()
//  vCurCursor.GoToStart(False)
//  Do
//    If NOT vCurCursor.gotoEndOfParagraph( True ) Then Exit Do
//    sCurStyle = vCurCursor.ParaStyleName
//    s = s & """" & sCurStyle & """" & CHR$(10)
//  Loop Until NOT vCurCursor.gotoNextParagraph( False )
//  MsgBox s, 0, "Styles in Document"
//End Sub

    public static void main(String[] args) {
        try {
            // get the remote office component context
            XComponentContext xContext = Bootstrap.bootstrap();

            XMultiComponentFactory xMCF = xContext.getServiceManager();

            XDesktop xDesktop =
              (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  xMCF.createInstanceWithContext("com.sun.star.frame.Desktop",xContext));


            XTextDocument xtextdoc = (XTextDocument)
                UnoRuntime.queryInterface(XTextDocument.class, xDesktop.getCurrentComponent());
            if (xtextdoc!=null){
                XText vText = xtextdoc.getText();
                XTextCursor vCurCursor = vText.createTextCursor();
                XParagraphCursor vParCursor= (XParagraphCursor) UnoRuntime.queryInterface(XParagraphCursor.class, vCurCursor);
                vCurCursor.gotoStart(false);
                String s="";
                while (!vParCursor.gotoNextParagraph(false)) {
                   XPropertySet parProp= (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, vParCursor);
                   s= s+"\""+parProp.getPropertyValue("ParaStypeName")+"\"\n";
                }
                System.out.println(s);
//    If NOT vCurCursor.gotoEndOfParagraph( True ) Then Exit Do
//    sCurStyle = vCurCursor.ParaStyleName
//    s = s & """" & sCurStyle & """" & CHR$(10)

            }else{
                System.out.println("You should have an open text document.");
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
