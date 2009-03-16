/*
 * Frysak_11_closing_document.java
 *
 * Created on 2009.02.06 - 11:19:49
 *
 * Close the current open document.
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.sdbc.XCloseable;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XModifiable;

/**
 *
 * @author brant
 */
public class Frysak_11_closing_document {
    
    /** Creates a new instance of Frysak_11_closing_document */
    public Frysak_11_closing_document() {
    }

//    -- try to get a script context, will be .nil, if script was not invoked by OOo
//    x_ScriptContext = uno.getScriptContext()
//    if (x_ScriptContext <> .nil) then
//    do
//        -- invoked by OOo as a macro
//        -- get context
//        x_ComponentContext = x_ScriptContext~getComponentContext
//        -- get desktop (an XDesktop)
//        x_Desktop = x_ScriptContext~getDesktop
//        -- get current document
//        x_Document = x_ScriptContext~getDocument
//    end
//    else
//    do
//        -- called from outside of OOo, create a connection
//        -- connect to Open Office and get component context
//        x_ComponentContext = UNO.connect()
//        -- create a desktop service and its interface
//        service = "com.sun.star.frame.Desktop"
//        s_Desktop = x_ComponentContext~getServiceManager~XMultiServiceFactory~createInstance(service)
//        x_Desktop = s_Desktop~XDesktop
//        -- get the last active document
//        x_Document = x_Desktop~getCurrentComponent()
//    end
//
//    -- create outputtext
//    output = "Document "
//    -- check if doucument has been changed
//    x_Modifiable = x_Document~XModifiable
//    x_Storable = x_Document~XStorable
//    If (x_Modifiable~isModified()) Then
//    do
//        output = output || "was Modified, "
//        /*
//        if there is allready a file containing the document
//        then save into this file, else just set modify flag to false
//        (do not save file)
//        */
//        If (x_Storable~hasLocation & (\ x_Storable~isReadOnly)) Then
//        do
//            x_Storable~store()
//            output = output || "and has been stored - "
//        end
//        else
//        do
//            x_Modifiable~setModified(.false)
//            output = output || "and has NOT been stored - "
//        end
//    end
//    else
//    do
//        output = output || "was NOT Modified - "
//    end
//    /*
//    next check for different methods to shut down
//    if we are able to create a XModel interface then also try to query a XClosable
//    interface to close document. If XCloseable interface is not available, use
//    Documents dispose method to close the document. If XModel interface query fails,
//    terminate the frame to shut down.
//    */
//    -- x_ServiceInfo = x_Document~XServiceInfo
//    -- If x_ServiceInfo~supportsService("com.sun.star.frame.XModel") then
//    -- I dont know why, but this does not work properly, therefore
//    x_Model = x_Document~XModel
//    if x_Model <> .nil then
//    do
//        x_Closeable = x_Document~XCloseable
//        If x_Closeable <> .nil then
//        do
//            x_Closeable~close(.true)
//            output = output || "closed by XClosable Interface (SOFTEST WAY)"
//        end
//        else
//        do
//            x_Document~dispose()
//            output = output || "closed by XDocument Interface (SOFT WAY)"
//        end
//    end
//    else
//    do
//        x_Desktop~terminate()
//        output = output || "closed by XDesktop Interface (HARD WAY)"
//    end
//    -- finaly show message what happened
//    .bsf.dialog~messageBox(output, "Closing Document...", "information")
//    ::requires UNO.CLS
//
    public static void main(String[] args) {
        try {
            XComponentContext context = Bootstrap.bootstrap();
            XMultiComponentFactory serviceManager = context.getServiceManager();
            XDesktop desktop =
              (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context));
            XComponent document= desktop.getCurrentComponent();
            String output= "Document";

            XModifiable modifiable= (XModifiable) UnoRuntime.queryInterface(XModifiable.class, document);
            XStorable storable= (XStorable) UnoRuntime.queryInterface(XStorable.class, document);
            if (modifiable.isModified()){
                output= output+" was Modified";
                if (storable.hasLocation() && storable.isReadonly()){
                    storable.store();
                    output = output + " and has been stored - ";
                }else{
                    modifiable.setModified(new Boolean(false));
                }
            }else{
                output = output + " was NOT Modified - ";
            }

            //    /*
            //    next check for different methods to shut down
            //    if we are able to create a XModel interface then also try to query a XClosable
            //    interface to close document. If XCloseable interface is not available, use
            //    Documents dispose method to close the document. If XModel interface query fails,
            //    terminate the frame to shut down.
            //    */
            XServiceInfo serviceInfo= (XServiceInfo) UnoRuntime.queryInterface(XServiceInfo.class, document);
            XModel model= (XModel) UnoRuntime.queryInterface(XModel.class, document);
            if (model!=null){
                XCloseable closable= (XCloseable) UnoRuntime.queryInterface(XCloseable.class, document);
                if (closable!=null){
                    closable.close();
                    output = output + " closed by XClosable Interface (SOFTEST WAY)";
                }else{
                    document.dispose();
                    output = output + " closed by XDocument Interface (SOFT WAY)";
                }
            }else{
                desktop.terminate();
                output = output + " closed by XDesktop Interface (HARD WAY)";
            }

            System.out.println(output);

        }
        catch (java.lang.Exception e){
            System.out.println("This macro expects an open document.");
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}
