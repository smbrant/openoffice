/*
 * Macro_09_header_footer.java
 *
 * Created on 2009.02.05 - 08:42:47
 *
 * Delete header and footer
 *
 */

package com.example;

import com.sun.star.beans.XPropertySet;
import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XNameAccess;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Macro_09_header_footer {
    
    /** Creates a new instance of Macro_09_header_footer */
    public Macro_09_header_footer() {
    }

// Original code from: http://codesnippets.services.openoffice.org/Writer/Writer.AddAndRemoveHeaderAndFooter.snip
//
//-- first we get access to the textdocument and its text object
//x_TextDocument = x_Document~XTextDocument
//x_Text = x_TextDocument~getText
//
//-- get the current current cursor
//s_CurrentController = x_TextDocument~getCurrentController()
//x_TextViewCursorSupplier = s_CurrentController~XTextViewCursorSupplier
//x_CurrentCursor = x_TextViewCursorSupplier~getViewCursor()
//
//-- get access to the properties of the current cursor and retrieve
//-- name of the current pagestyle
//cursorproperties = x_CurrentCursor~XPropertySet
//pagestylename = cursorproperties~getPropertyValue("PageStyleName")
//
//-- next we search the StyleFamily entries for our current pagestyle
//x_StyleFamiliesSupplier = x_Document~XStyleFamiliesSupplier
//x_StyleFamilies = x_StyleFamiliesSupplier~getStyleFamilies()
//s_StyleFamily = x_StyleFamilies~getByName("PageStyles")
//x_NameAccess = s_StyleFamily~XNameAccess
//s_PageProperties = x_NameAccess~getByName(pagestylename)
//
//-- get the properties of the current page
//pageproperties = s_PageProperties~XPropertySet
//
//-- get the current status of header and footer of this page
//oldheader = pageproperties~getPropertyValue("HeaderIsOn")
//oldfooter = pageproperties~getPropertyValue("FooterIsOn")
//
//-- if header is on turn it of and vice versa
//-- i dont know why disabling needs a complet object as parameter, but
//-- enabling does not. (enabling wont work if using objects)
//
//if oldheader then pageproperties~setPropertyValue("HeaderIsOn", .bsf~new("java.lang.Boolean", 0))
//else pageproperties~setPropertyValue("HeaderIsOn", 1)
//
//if oldfooter then pageproperties~setPropertyValue("FooterIsOn", .bsf~new("java.lang.Boolean", 0))
//else pageproperties~setPropertyValue("FooterIsOn", 1)

    public static void main(String[] args) {
        try {
            XComponentContext context = Bootstrap.bootstrap();
            XMultiComponentFactory serviceManager = context.getServiceManager();
            XDesktop desktop =
              (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context));
            XTextDocument textDocument = (XTextDocument)
                UnoRuntime.queryInterface(XTextDocument.class, desktop.getCurrentComponent());
            if (textDocument==null){
                System.out.println("Macro expects an open document.");
            }else{
                XText text = textDocument.getText();
                XTextCursor textCursor = text.createTextCursor();
                XPropertySet cursorProperties= (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, textCursor);
                String pageStyleName= (String) cursorProperties.getPropertyValue("PageStyleName");
                XStyleFamiliesSupplier styleFamiliesSupplier= (XStyleFamiliesSupplier) UnoRuntime.queryInterface(XStyleFamiliesSupplier.class, textDocument);
                XNameAccess stylesFamilies= styleFamiliesSupplier.getStyleFamilies();

//                Object styleFamily=  stylesFamilies.getByName("PageStyles");
//                XNameAccess nameAccess= (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, styleFamily);
//                XPropertySet pageProperties= (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class,
//                        nameAccess.getByName(pageStyleName) );

                XNameAccess styleFamily= (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class,
                        stylesFamilies.getByName("PageStyles"));
                XPropertySet pageProperties= (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class,
                        styleFamily.getByName(pageStyleName) );

                if((Boolean)pageProperties.getPropertyValue("HeaderIsOn")){
                    pageProperties.setPropertyValue("HeaderIsOn", new Boolean(false));
                }
                if((Boolean)pageProperties.getPropertyValue("FooterIsOn")){
                    pageProperties.setPropertyValue("FooterIsOn", new Boolean(false));
                }

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
