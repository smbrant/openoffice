package com.example;

import com.sun.star.awt.PushButtonType;
import com.sun.star.awt.XTextComponent;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.lib.uno.helper.Factory;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.ucb.XSimpleFileAccess;
import com.sun.star.ui.dialogs.XFilePicker;


public final class AddOnMacros extends WeakBase
   implements com.sun.star.lang.XServiceInfo,
              com.sun.star.frame.XDispatchProvider,
              com.sun.star.lang.XInitialization,
              com.sun.star.frame.XDispatch
{
    private final XComponentContext m_xContext;
    private com.sun.star.frame.XFrame m_xFrame;
    private static final String m_implementationName = AddOnMacros.class.getName();
    private static final String[] m_serviceNames = {
        "com.sun.star.frame.ProtocolHandler" };


    public AddOnMacros( XComponentContext context )
    {
        m_xContext = context;
    };

    public static XSingleComponentFactory __getComponentFactory( String sImplementationName ) {
        XSingleComponentFactory xFactory = null;

        if ( sImplementationName.equals( m_implementationName ) )
            xFactory = Factory.createComponentFactory(AddOnMacros.class, m_serviceNames);
        return xFactory;
    }

    public static boolean __writeRegistryServiceInfo( XRegistryKey xRegistryKey ) {
        return Factory.writeRegistryServiceInfo(m_implementationName,
                                                m_serviceNames,
                                                xRegistryKey);
    }

    // com.sun.star.lang.XServiceInfo:
    public String getImplementationName() {
         return m_implementationName;
    }

    public boolean supportsService( String sService ) {
        int len = m_serviceNames.length;

        for( int i=0; i < len; i++) {
            if (sService.equals(m_serviceNames[i]))
                return true;
        }
        return false;
    }

    public String[] getSupportedServiceNames() {
        return m_serviceNames;
    }

    // com.sun.star.frame.XDispatchProvider:
    public com.sun.star.frame.XDispatch queryDispatch( com.sun.star.util.URL aURL,
                                                       String sTargetFrameName,
                                                       int iSearchFlags )
    {
        if ( aURL.Protocol.compareTo("com.example.addonmacros:") == 0 )
        {
            if ( aURL.Path.compareTo("Command1") == 0 )
                return this;
            if ( aURL.Path.compareTo("Command2") == 0 )
                return this;
            if ( aURL.Path.compareTo("Command3") == 0 )
                return this;
            if ( aURL.Path.compareTo("Command4") == 0 )
                return this;
            if ( aURL.Path.compareTo("Command5") == 0 )
                return this;
            if ( aURL.Path.compareTo("Command6") == 0 )
                return this;
            if ( aURL.Path.compareTo("Command7") == 0 )
                return this;
            if ( aURL.Path.compareTo("Command8") == 0 )
                return this;
            if ( aURL.Path.compareTo("Command9") == 0 )
                return this;
        }
        return null;
    }

    // com.sun.star.frame.XDispatchProvider:
    public com.sun.star.frame.XDispatch[] queryDispatches(
         com.sun.star.frame.DispatchDescriptor[] seqDescriptors )
    {
        int nCount = seqDescriptors.length;
        com.sun.star.frame.XDispatch[] seqDispatcher =
            new com.sun.star.frame.XDispatch[seqDescriptors.length];

        for( int i=0; i < nCount; ++i )
        {
            seqDispatcher[i] = queryDispatch(seqDescriptors[i].FeatureURL,
                                             seqDescriptors[i].FrameName,
                                             seqDescriptors[i].SearchFlags );
        }
        return seqDispatcher;
    }

    // com.sun.star.lang.XInitialization:
    public void initialize( Object[] object )
        throws com.sun.star.uno.Exception
    {
        if ( object.length > 0 )
        {
            m_xFrame = (com.sun.star.frame.XFrame)UnoRuntime.queryInterface(
                com.sun.star.frame.XFrame.class, object[0]);
        }
    }

    // com.sun.star.frame.XDispatch:
    public void dispatch( com.sun.star.util.URL aURL,
                           com.sun.star.beans.PropertyValue[] aArguments )
    {
         if ( aURL.Protocol.compareTo("com.example.addonmacros:") == 0 )
        {
            // Hello
            //============================================================================================
            if ( aURL.Path.compareTo("Command1") == 0 ){
                // Definição do texto do documento
                XTextDocument xDoc = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class,
                     m_xFrame.getController().getModel());
                XText text = xDoc.getText();
                XTextCursor textCursor = text.createTextCursor();
                text.insertString(textCursor, "Hello from Macro1!"+"\n\n",false);
                return;
            }

            // Clear blank lines
            //============================================================================================
            if ( aURL.Path.compareTo("Command2") == 0 )
                try {
                    Writer writer= new Writer(m_xContext);
                    writer.clearBlankLines();
                }
                catch (java.lang.Exception e){e.printStackTrace();}
                finally {return;}

            // Show open file dialog
            //============================================================================================
            if ( aURL.Path.compareTo("Command3") == 0 )
                try {
                    // Diálogo de abertura de arquivo
                    XMultiComponentFactory srvMan = m_xContext.getServiceManager();
                    XFilePicker vFileDialog=
                        (XFilePicker) UnoRuntime.queryInterface(XFilePicker.class,
                        srvMan.createInstanceWithContext("com.sun.star.ui.dialogs.FilePicker", m_xContext));
                    XSimpleFileAccess vFileAccess=
                            (XSimpleFileAccess) UnoRuntime.queryInterface(XSimpleFileAccess.class,
                            srvMan.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", m_xContext));
                    String sInitPath = "c:\\temp";
                    if (vFileAccess.exists(sInitPath)){
                        vFileDialog.setDisplayDirectory(sInitPath);
                    }
                    short iAccept= vFileDialog.execute();
                    if (iAccept == 1){
                        String[] files= vFileDialog.getFiles();
                        System.out.println(files[0]);
                    }
                }
                catch (java.lang.Exception e){ e.printStackTrace(); }
                finally {return;}

            // Dialog asking for data to be inserted in document
            //============================================================================================
            if ( aURL.Path.compareTo("Command4") == 0 )
                // Open a dialog and insert the informed text in document
                try {
                    Dialog dialog = new Dialog(m_xContext,200,200,80,280,"Send Mail","sendMailDialog");

                    dialog.insertFixedText(10, 10, 30, "Subject:");
                    XTextComponent subjectField = dialog.insertTextField(42, 10, 228, "Subject...");
                    dialog.insertFixedText(10, 30, 30, "Recipients:");
                    XTextComponent recipientsField = dialog.insertTextField(42, 30, 228, "Recipients...");

                    dialog.insertButton(105, 60, 30, "OK", PushButtonType.OK_value);
                    dialog.insertButton(145, 60, 30, "Cancel", PushButtonType.CANCEL_value);

                    short returnValue = dialog.execute();

                    XMultiComponentFactory xMCF = m_xContext.getServiceManager();
                    XDesktop xDesktop =
                      (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                          xMCF.createInstanceWithContext("com.sun.star.frame.Desktop",m_xContext));

                    XTextDocument xtextdoc = (XTextDocument)
                        UnoRuntime.queryInterface(XTextDocument.class, xDesktop.getCurrentComponent());
                    if (xtextdoc!=null){
                        if (returnValue==1){
                            XText xtext = xtextdoc.getText();
                            XTextCursor xtextcursor = xtext.createTextCursor();
                            xtext.insertString(xtextcursor, "Subject: "+subjectField.getText()+"\n"+
                                                            "Recipients: "+recipientsField.getText()+"\n\n",
                                    false);
                        }
                    }
                }
                catch (java.lang.Exception e){ e.printStackTrace(); }
                finally {return;}

            // Clear header and footer
            //============================================================================================
            if ( aURL.Path.compareTo("Command5") == 0 )
                try {
                    XComponentContext context = m_xContext;
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
                catch (java.lang.Exception e){ e.printStackTrace(); }
                finally {return;}

            //
            // Set background color
            //
            // Adapted from http://codesnippets.services.openoffice.org/Writer/Writer.ChangeDocumentBackgroundColor.snip
            //============================================================================================
            if ( aURL.Path.compareTo("Command6") == 0 )
                try{
                    XComponentContext context = m_xContext;
                    XMultiComponentFactory serviceManager = context.getServiceManager();
                    XDesktop desktop =
                      (XDesktop) UnoRuntime.queryInterface( XDesktop.class,
                          serviceManager.createInstanceWithContext("com.sun.star.frame.Desktop",context));
                    XTextDocument oDoc = (XTextDocument)
                        UnoRuntime.queryInterface(XTextDocument.class, desktop.getCurrentComponent());

                    // create a supplier to get the Style family collection
                    XStyleFamiliesSupplier xSupplier = ( XStyleFamiliesSupplier ) UnoRuntime.queryInterface(XStyleFamiliesSupplier.class, oDoc );

                    // get the NameAccess interface from the Style family collection
                    XNameAccess xNameAccess = xSupplier.getStyleFamilies();

                    XNameContainer xPageStyleCollection = (XNameContainer) UnoRuntime.queryInterface(XNameContainer.class, xNameAccess.getByName( "PageStyles" ));

                    // create a PropertySet to set the properties for the new Pagestyle
                    XPropertySet xPropertySet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xPageStyleCollection.getByName("Standard") );

                    // setBackgroundColor
                    xPropertySet.setPropertyValue("BackColor",new Integer( (int)255 ) );
                }
                catch (java.lang.Exception e){ e.printStackTrace();} finally {return;}

            //
            //
            //
            //
            //============================================================================================
            if ( aURL.Path.compareTo("Command7") == 0 )
            {
                // add your own code here
                return;
            }
            if ( aURL.Path.compareTo("Command8") == 0 )
            {
                // add your own code here
                return;
            }
            if ( aURL.Path.compareTo("Command9") == 0 )
            {
                // add your own code here
                return;
            }
        }
    }

    public void addStatusListener( com.sun.star.frame.XStatusListener xControl,
                                    com.sun.star.util.URL aURL )
    {
        // add your own code here
    }

    public void removeStatusListener( com.sun.star.frame.XStatusListener xControl,
                                       com.sun.star.util.URL aURL )
    {
        // add your own code here
    }

}
