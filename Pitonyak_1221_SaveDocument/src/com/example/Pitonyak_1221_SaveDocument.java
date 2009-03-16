/*
 * Pitonyak_1221_SaveDocument.java
 *
 * Created on 2009.02.17 - 08:40:11
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XModifiable;

/**
 *
 * @author brant
 */
public class Pitonyak_1221_SaveDocument {
    
    /** Creates a new instance of Pitonyak_1221_SaveDocument */
    public Pitonyak_1221_SaveDocument() {
    }

    //Proper method to save a document:
    //
    //If (ThisComponent.isModified()) Then
    //    If (ThisComponent.hasLocation() AND (Not ThisComponent.isReadOnly())) Then
    //        ThisComponent.store()
    //    Else
    //        REM Either the document does not have a location or you cannot
    //        REM save the document because the location is read-only.
    //        setModified(False)
    //    End If
    //End If
    public static void main(String[] args) {
        try {
            XComponentContext context = Bootstrap.bootstrap();
            XMultiComponentFactory serviceManager= context.getServiceManager();
            XDesktop desktop= (XDesktop) UnoRuntime.queryInterface(XDesktop.class,
                    serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context));
            XComponent document= desktop.getCurrentComponent();
            XModifiable modifiable= (XModifiable) UnoRuntime.queryInterface(XModifiable.class, document);
            if (document!=null)
                if (modifiable.isModified()) {
                    XStorable store= (XStorable) UnoRuntime.queryInterface(XStorable.class, document);
                    if (store.hasLocation() && store.isReadonly()){
                        store.store();
                    }else{
                        System.out.println("Either the document does not have a location or you cannot");
                        System.out.println("save the document because the location is read-only.");
                        modifiable.setModified(false);  // why? This is not good because you will
                                                        // be able to close modified doc. without warning.
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
