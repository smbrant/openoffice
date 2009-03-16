/*
 * Pitonyak_09_05.java
 *
 * Created on 2009.02.03 - 10:00:04
 *
 * Testing if a service supports an interface
 *
 */

package com.example;

import com.sun.star.beans.PropertyValue;
import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sdbc.XStruct;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XInterface;

/**
 *
 * @author brant
 */
public class Pitonyak_09_05 {
    
    /** Creates a new instance of Pitonyak_09_05 */
    public Pitonyak_09_05() {
    }
    
//Dim aProp As New com.sun.star.beans.PropertyValue
//Print IsUnoStruct(aProp) 'True
//Print HasUnoInterfaces(aProp, "com.sun.star.uno.XInterface") 'False
//Print IsUnoStruct(ThisComponent) 'False
//Print HasUnoInterfaces(ThisComponent, "com.sun.star.uno.XInterface") 'True
        public static void main(String[] args) {
        try {
            XComponentContext context = Bootstrap.bootstrap();
            XMultiComponentFactory srvMan = context.getServiceManager();

            PropertyValue aProp= new PropertyValue();

            XStruct xXStruct= (XStruct)UnoRuntime.queryInterface(XInterface.class, aProp);
            System.out.println(xXStruct); //not null

            XInterface xInterface= (XInterface)UnoRuntime.queryInterface(XInterface.class, aProp);
            System.out.println(xInterface); //null

            XDesktop xDesktop =
              (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  srvMan.createInstanceWithContext("com.sun.star.frame.Desktop",context));

            XTextDocument textDoc = (XTextDocument)
                UnoRuntime.queryInterface(XTextDocument.class, xDesktop.getCurrentComponent());

            xInterface= (XInterface)UnoRuntime.queryInterface(XInterface.class, textDoc);
            System.out.println(xInterface); // not null if doc opened (a object implementing the interface)

        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
