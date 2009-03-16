/*
 * Pitonyak_09_04.java
 *
 * Created on 2009.02.03 - 08:44:37
 *
 * Component properties
 *
 */

package com.example;

import com.sun.star.beans.Property;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XNameAccess;
import com.sun.star.document.XDocumentInfo;
import com.sun.star.document.XDocumentInfoSupplier;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextTablesSupplier;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import javax.swing.JOptionPane;

/**
 *
 * @author brant
 */
public class Pitonyak_09_04 {
    
    /** Creates a new instance of Pitonyak_09_04 */
    public Pitonyak_09_04() {
    }

    //ch09 lis04
    //Dim v
    //v = Thiscomponent.getTextTables()
    //Print IsObject(v) 'True
    //Print IsNull(v) 'False
    //Print IsEmpty(v) 'False
    //Print IsArray(v) 'False
    //Print IsUnoStruct(v) 'False
    //Print TypeName(v) 'Object
    //MsgBox v.dbg_methods 'This property is discussed later
    public static void main(String[] args) throws Exception {
        XComponentContext context = Bootstrap.bootstrap();
        XMultiComponentFactory srvMan = context.getServiceManager();
        XDesktop xDesktop =
          (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
              srvMan.createInstanceWithContext("com.sun.star.frame.Desktop",context));
        XTextDocument textDoc = (XTextDocument)
            UnoRuntime.queryInterface(XTextDocument.class, xDesktop.getCurrentComponent());

        XTextTablesSupplier textTables= (XTextTablesSupplier)
            UnoRuntime.queryInterface(XTextTablesSupplier.class, textDoc);

        try{  
            XNameAccess v= textTables.getTextTables();
            System.out.println(v);
            //TODO: how to show msgbox in java
        }
        catch (java.lang.Exception e){
            System.out.println("You should have an open document.");
        }


        System.exit( 0 );
    }
    
}
