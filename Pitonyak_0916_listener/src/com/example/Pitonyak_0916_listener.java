/*
 * Pitonyak_0916_listener.java
 *
 * Created on 2009.02.11 - 11:55:51
 *
 */

package com.example;

import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.document.XEventListener;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;

/**
 *
 * @author brant
 */
public class Pitonyak_0916_listener {
    
    /** Creates a new instance of Pitonyak_0916_listener */
    public Pitonyak_0916_listener() {
    }

    //Sub MyFirstListener
    //    Dim vListener
    //    Dim vEventObj As New com.sun.star.lang.EventObject
    //    Dim sPrefix$
    //    Dim sService$
    //    sPrefix = "first_listen_"
    //    sService = "com.sun.star.lang.XEventListener"
    //    vListener = CreateUnoListener(sPrefix, sService)
    //    vListener.disposing(vEventObj)
    //End Sub
    public static void main(String[] args) {
        try {
            XComponentContext context = Bootstrap.bootstrap();
            XMultiComponentFactory serviceManager = context.getServiceManager();
            XDesktop desktop= (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                  serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context));

            XComponent document= desktop.getCurrentComponent();

            XEventListener listener= (XEventListener) UnoRuntime.queryInterface( XEventListener.class,
                  serviceManager.createInstanceWithContext("com.sun.star.lang.XEventListener",context));

//          document.addEventListener(arg0)
//
//          EventObject eventObj= new EventObject();

//            System.out.println(listener);

//            service.
//            listener.disposing(arg0);



        }
        catch (java.lang.Exception e){
            e.printStackTrace();
        }
        finally {
            System.exit( 0 );
        }
    }
    
}

// from http://osdir.com/ml/openoffice.devel.api/2007-01/msg00062.html

//import java.awt.Frame;
//import java.net.MalformedURLException;
//
//import org.eclipse.swt.widgets.Shell;
//
//import com.sun.star.beans.Property;
//import com.sun.star.beans.PropertyValue;
//import com.sun.star.beans.XPropertySet;
//import com.sun.star.comp.beans.OOoBean;
//import com.sun.star.document.XEventBroadcaster;
//import com.sun.star.uno.UnoRuntime;
//import com.sun.star.uno.XComponentContext;
//
//
//public class TestConnection {
//
//static com.sun.star.uno.XComponentContext xContext = null;
//static long SLEEPTIME = 3000;
///**
//* @param args
//*/
//public static void main(String[] args) {
//try {
//// get the remote office component context
//// xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();
//
//// System.out.println("Connected to a running office ...");
//
//Shell oShell = new Shell();
//Frame oFrame = new Frame();
//System.out.println("Creating bean");
//OOoBean oBean = new OOoBean();
//System.out.println("Created Bean");
//if (oBean.getOOoConnection()!=null)
//System.out.println("Connection is not null");
//else
//System.out.println("Connection is null");
//oFrame.add(oBean);
//oFrame.setVisible(true);
//oFrame.setSize(500, 400);
//
////registro il listener per gli eventi sul controllo
//office
//XComponentContext xRemoteContext = oBean.getOOoConnection().getComponentContext();
//String st2[]=xRemoteContext.getServiceManager().getAvailableServiceNames();
//System.out.println(st2);
//Object xGlobalBroadCaster = xRemoteContext.getServiceManager().createInstanceWithContext(
//"com.sun.star.frame.GlobalEventBroadcaster",
//xRemoteContext);
//XEventBroadcaster xEventBroad = (XEventBroadcaster)UnoRuntime.queryInterface(XEventBroadcaster.class, xGlobalBroadCaster);
//xEventBroad.addEventListener(new
//com.sun.star.document.XEventListener() {
//public void
//notifyEvent(com.sun.star.document.EventObject oEvent) {
//// the control model which fired the event
//System.out.println("Evento da : " + oEvent.Source + " " + oEvent.EventName);
//XPropertySet xSourceProps = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class,oEvent.Source);
//if (xSourceProps!=null) {
//// Property[] aoProps = xSourceProps.getPropertySetInfo().getProperties();
//// for(int iCount=0;
//iCount<aoProps.length; iCount++) {
//// System.out.println("Property " + iCount + ": " + aoProps[iCount].Name);
//// }
////String sUid =
//(String)xSourceProps.getPropertyValue("RuntimeUID");
//}
//}
//public void
//disposing(com.sun.star.lang.EventObject e) {
//System.out.println("On Dispose");
//}
//});
