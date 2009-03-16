// from http://www.openoffice.org/servlets/ReadMsg?list=allcvs&msgNo=67637
package tests;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.awt.FontWeight;

import com.sun.star.beans.PropertyState;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertyState;

import com.sun.star.bridge.XUnoUrlResolver;

import com.sun.star.comp.servicemanager.ServiceManager;

import com.sun.star.connection.XConnector;
import com.sun.star.connection.XConnection;

import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.container.XNamed;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XIndexReplace;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;

import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapeGrouper;
import com.sun.star.drawing.XShapes;
import com.sun.star.drawing.XDrawPageSupplier;

import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XComponentLoader;

import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceInfo;

import com.sun.star.style.NumberingType;
import com.sun.star.style.XStyle;
import com.sun.star.style.XStyleFamiliesSupplier;

import com.sun.star.table.XCell;

import com.sun.star.text.ControlCharacter;
import com.sun.star.text.ReferenceFieldSource;
import com.sun.star.text.ReferenceFieldPart;
import com.sun.star.text.TextColumn;
import com.sun.star.text.TextContentAnchorType;
import com.sun.star.text.XAutoTextContainer;
import com.sun.star.text.XAutoTextGroup;
import com.sun.star.text.XAutoTextEntry;
import com.sun.star.text.XDependentTextField;
import com.sun.star.text.XDocumentIndex;
import com.sun.star.text.XFootnote;
import com.sun.star.text.XFootnotesSupplier;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XReferenceMarksSupplier;
import com.sun.star.text.XRelativeTextContentInsert;
import com.sun.star.text.XSentenceCursor;
import com.sun.star.text.XSimpleText;
import com.sun.star.text.XText;
import com.sun.star.text.XTextColumns;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextField;
import com.sun.star.text.XTextFrame;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextSection;
import com.sun.star.text.XTextTable;
import com.sun.star.text.XTextTableCursor;
import com.sun.star.text.XTextTablesSupplier;
import com.sun.star.text.XWordCursor;

import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XInterface;
import com.sun.star.uno.XNamingService;

import com.sun.star.util.XRefreshable;

import java.lang.Math;
import java.util.Random;

/** Writer examples
@author Martin Gallwey
*/

/* These examples show how to use the following features of the Writer API:

o Text Documents
o Text Fields
o Paragraphs
o Text Frames
o Text Tables
o Text Sections
o Text Columns
o Numbering Ruels
o Styles
o Indices
o References
o Foot notes
o Shapes and the draw page
o Auto Texts

 Each example is in a separate method call.

 It accepts one command line parameter in the form of a connection string,
 if no connection string is specified it uses the default of
 "socket,host=localhost,port=8100"
*/
public class WriterExamples
{
    // Make commonly used interfaces available to all methods

    // The MultiServiceFactory interface of the text document
    protected XMultiServiceFactory mxDocFactory = null;
    // The MultiServiceFactory interface of the Office
    protected XMultiServiceFactory mxFactory = null;
    // The text document's property set
    protected XPropertySet  mxDocProps = null;
    // The text document's XText interface for inserting or
    // removing XTextContents
    protected XText         mxDocText = null;
    // The text document's cursor for traversing the document
    protected XTextCursor   mxDocCursor = null;
    // The text document itself, we will often use queryInterface
    // on it to access other interfaces that it supports such
    // as XDrawPageSupplier to get access to the draw page
    protected XTextDocument mxDoc = null;
    // We'll keep a reference to the fish section in case
    // we need to use the XRelativeTextContentInsert interface
    protected XTextContent 	mxFishSection = null;
    // A random number generator
    protected Random 		maRandom = null;

    public static void main(String args[])
    {
        String sConnection;
        // Get the connect string, defaulting to localhost, port 8100
        // if non supplied as command line arguments
        if( args.length == 0 )
        {
            sConnection = "socket,host=localhost,port=8100";
            System.out.println( "Using default connect string: " + sConnection );
        }
        else
        {
            sConnection = args[0];
            System.out.println( "Using connect string: " + sConnection );
        }

        // Create an instance of the class and call it's begin method
        WriterExamples aExample = new WriterExamples();
        aExample.begin ( sConnection );
    }
    public void begin(String sConnection)
    {
        // Call connect, after this call mxFactory should refer to
        // the MultiServiceFactory of a running office. If something
        // goes wrong, print a stack trace and exit.
        try
        {
            connect( sConnection );
        }
        catch( Exception e )
        {
            e.printStackTrace ( System.out );
            return;
        }

        // Call GetDocument to set mxDoc to refer to a new, empty, writer document
        System.out.println("Opening an empty Writer document");
        if ( GetDocument() == false )
        {
            // If we couldn't get one, exit as we can't continue without one!
            System.out.println("Couldn't get an empty Writer document! Exiting.");
            return;
        }

        // get a new random generator
        maRandom = new Random();
        // get a reference to a commonly used interface
        mxDocText = mxDoc.getText();

        // Get a reference to the document's property set. This contains document
        // information like the current word count
        mxDocProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, mxDoc );

        // Simple text insertion example
        BodyTextExample ();
        // Example using text ranges to insert strings at the beginning or end of a text range
        TextRangeExample ();

        // Create a document cursor and remember it, it will be used in most of the following
        // examples
        mxDocCursor = mxDocText.createTextCursor();

        // Demonstrate some of the different cursor types (word, sentence)
        TextCursorExample ();

        // Access the text document's multi service factory, which we will need for most of the
        // following examples
        mxDocFactory = (XMultiServiceFactory) UnoRuntime.queryInterface( XMultiServiceFactory.class, mxDoc );

        // Examples of text fields, dependant text fields and field masters
        TextFieldExample ();

        // Example of using an XEnumerationAccess to iterate over paragraphs and set properties of each
        // paragraph as we do so
        ParagraphExample ();

        // Example of creating and manipulating a text frame
        TextFrameExample ();

        // Example of creating and manipulating a text table, text table rows and text table cells
        TextTableExample ();

        // Example of creating, inserting and manipulating text sections, as well as an example
        // of how to refresh the document
        TextSectionExample ();

        // Example of creating a text section over a block of text and formatting the text section
        // into columns, as well as how to insert an empty paragraph using the XRelativeTextContentInsert
        TextColumnsExample ();

        // Example of creating the NumberingRules service and adjusting NumberingTypes and
        // NumberingLevels
        NumberingExample ();

        // Example of how to use the XStyleFamiliesSupplier interface of the document and how to
        // create, insert and apply styles
        StylesExample ();
        IndexExample ();

        // Example of how to create and manipulate reference marks and GetReference text fields
        ReferenceExample ();

        // Example of how to create and insert Footnotes and how to use the XFootnotesSupplier
        // interface of the document
        FootnoteExample ();

        // This method demonstrates how to create shapes from the document factory and how to
        // access the draw page of the document using the XDrawPageSupplier interface
        DrawPageExample ();


        // This example demonstrates the use of the AutoTextContainer, AutoTextGroup and AutoTextEntry services
        // and shows how to create, insert and modify auto text blocks
        AutoTextExample ();
        System.exit ( 0 );
    }

    /** Connect to a specified host which is running office and accepting
        connections on the requested port. Set up our class variable 'mxFactory' to
        reference the ServiceManager of the office */

    public void connect( String sConnection )
        throws com.sun.star.uno.Exception,
        com.sun.star.uno.RuntimeException, Exception
    {
        XComponentContext xContext =
            com.sun.star.comp.helper.Bootstrap.createInitialComponentContext( null );
        XMultiComponentFactory xLocalServiceManager = xContext.getServiceManager();

        Object  xUrlResolver  = xLocalServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", xContext );
        XUnoUrlResolver urlResolver = (XUnoUrlResolver)UnoRuntime.queryInterface( XUnoUrlResolver.class, xUrlResolver );
        Object rInitialObject = urlResolver.resolve( "uno:" + sConnection + ";urp;StarOffice.NamingService" );
        XNamingService rName = (XNamingService)UnoRuntime.queryInterface(XNamingService.class, rInitialObject );
        if( rName != null )
        {
            Object rXsmgr = rName.getRegisteredObject("StarOffice.ServiceManager" );
            mxFactory = (XMultiServiceFactory)
                UnoRuntime.queryInterface( XMultiServiceFactory.class, rXsmgr );
        }
    }

    /** Use the Office's MultiServiceFactory to create a new, empty, writer document.
        Set our 'MxDoc' class variable to refer to this new document
     */
    public boolean GetDocument ()
    {
        try
        {
            // Use the Office's MultiServiceFactory to create the 'com.sun.star.frame.Desktop' service,
            // and then access it's XComponentLoader interface
            XComponentLoader xLoader = ( XComponentLoader ) UnoRuntime.queryInterface( XComponentLoader.class,
                        mxFactory.createInstance( "com.sun.star.frame.Desktop" ) );
            // Create an empty array of PropertyValues as loadComponentFromURL expects to be passed one
            PropertyValue [] aEmptyArgs = new PropertyValue [0];
            // Use the component loader to load a new, empty, writer document
            mxDoc = (XTextDocument) UnoRuntime.queryInterface ( XTextDocument.class,
                       xLoader.loadComponentFromURL( "private:factory/swriter", "_blank", 0, aEmptyArgs ) );
        }
        catch( Exception e )
        {
            // If anything goes wrong, give the user a stack trace
            e.printStackTrace ( System.out );
            // And return false to notify the user that we couldn't get a document...and can't
            // continue
            return false;
        }
        return true;
    }

    protected void BodyTextExample ()
    {
        // Body Text and TextDocument example
        try
        {
            // demonstrate simple text insertion
            mxDocText.setString ( "This is the new body text of the document.\n\nThis is on the second line.\n\n" );
        }
        catch ( Exception e )
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "XBodyText/XTextDocument example complete. Word count of document is now " + GetWordCount() );
    }
    /** This method demonstrates the XTextRange interface
     */
    protected void TextRangeExample ()
    {
        try
        {
            // Get a text range refering to the beginning of the text document
            XTextRange xStart = mxDocText.getStart();
            // use setString to insert text at the beginning
            xStart.setString ( "This is text inserted at the beginning.\n\n" );
            // Get a text range refering to the end of the text document
            XTextRange xEnd = mxDocText.getEnd();
            // use setString to insert text at the end
            xEnd.setString ( "This is text inserted at the end.\n\n" );
        }
        catch ( Exception e )
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "XTextRange example complete. Word count of document is now " + GetWordCount() );
    }
    /** This method demonstrates how to access the different types of text cursor, and gives
        examples on their usage
     */
    protected void TextCursorExample ()
    {
        try
        {
            // First, get the XSentenceCursor interface of our text cursor
            XSentenceCursor xSentenceCursor = (XSentenceCursor) UnoRuntime.queryInterface( XSentenceCursor.class, mxDocCursor );
            // Goto the next cursor, without selecting it
            xSentenceCursor.gotoNextSentence( false );
            // Get the XWordCursor interface of our text cursor
            XWordCursor xWordCursor = (XWordCursor) UnoRuntime.queryInterface( XWordCursor.class, mxDocCursor );
            // Skip the first four words of this sentence and select the fifth
            xWordCursor.gotoNextWord( false );
            xWordCursor.gotoNextWord( false );
            xWordCursor.gotoNextWord( false );
            xWordCursor.gotoNextWord( false );
            xWordCursor.gotoNextWord( true );

            // Use the XSimpleText interface to insert a word at the current cursor location, over-writing
            // the current selection (the fifth word selected above)
            mxDocText.insertString ( xWordCursor, "old ", true );

            // Access the property set of the cursor, and set the currently selected text (which is
            // the string we just inserted) to be bold
            XPropertySet xCursorProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, mxDocCursor );
            xCursorProps.setPropertyValue ( "CharWeight", new Float ( com.sun.star.awt.FontWeight.BOLD ) );

            // replace the '.' at the end of the sentence with a new string
            xSentenceCursor.gotoEndOfSentence( false );
            xWordCursor.gotoPreviousWord( true );
            mxDocText.insertString ( xWordCursor, ", which has been changed with text cursors!", true);
        }
        catch ( Exception e )
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "XTextCursor example complete. Word count of document is now " + GetWordCount() );
    }
    /** This method inserts both a date field and a user field containing the number '42'
     */
    protected void TextFieldExample ()
    {
        try
        {
            // Use the text document's factory to create a DateTime text field, and access it's
            // XTextField interface
            XTextField xDateField = (XTextField) UnoRuntime.queryInterface ( XTextField.class,
                                     mxDocFactory.createInstance ( "com.sun.star.text.TextField.DateTime" ) );
            // Insert it at the end of the document
            mxDocText.insertTextContent ( mxDocText.getEnd(), xDateField, false );
            // Use the text document's factory to create a user text field, and access it's
            // XDependantTextField interface
            XDependentTextField xUserField = (XDependentTextField) UnoRuntime.queryInterface ( XDependentTextField.class,
                                     mxDocFactory.createInstance ( "com.sun.star.text.TextField.User" ) );
            // Create a fieldmaster for our newly created User Text field, and access it's
            // XPropertySet interface
            XPropertySet xMasterPropSet = (XPropertySet) UnoRuntime.queryInterface ( XPropertySet.class,
                                     mxDocFactory.createInstance ( "com.sun.star.text.FieldMaster.User" ) );
            // Set the name and value of the FieldMaster
            xMasterPropSet.setPropertyValue ( "Name", "UserEmperor" );
            xMasterPropSet.setPropertyValue ( "Value", new Integer ( 42 ) );
            // Attach the field master to the user field
            xUserField.attachTextFieldMaster ( xMasterPropSet );
            // Move the cursor to the end of the document
            mxDocCursor.gotoEnd( false );
            // insert a paragraph break using the XSimpleText interface
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );
            // Insert the user field at the end of the document
            mxDocText.insertTextContent ( mxDocText.getEnd(), xUserField, false );
        }
        catch ( Exception e )
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "XTextField example complete. Word count of document is now " + GetWordCount() );
    }
    protected void ParagraphExample ()
    {
        try
        {
            // The service 'com.sun.star.text.Text' supports the XEnumerationAccess interface to provide an enumeration
            // of the paragraphs contained by the text the service refers to.

            // Here, we access this interface
            XEnumerationAccess xParaAccess = (XEnumerationAccess) UnoRuntime.queryInterface( XEnumerationAccess.class, mxDocText );
            // Call the XEnumerationAccess's only method to access the actual Enumeration
            XEnumeration xParaEnum = xParaAccess.createEnumeration();

            // While there are paragraphs, do things to them
            while ( xParaEnum.hasMoreElements() )
            {
                // Get a reference to the next paragraphs XServiceInfo interface. TextTables are also part of this
                // enumeration access, so we ask the element if it is a TextTable, if it doesn't support the
                // com.sun.star.text.TextTable service, then it is safe to assume that it really is a paragraph
                XServiceInfo xInfo = (XServiceInfo) UnoRuntime.queryInterface( XServiceInfo.class, xParaEnum.nextElement() );
                if ( !xInfo.supportsService ( "com.sun.star.text.TextTable" ) )
                {
                    // Access the paragraph's property set...the properties in this property set are listed
                    // in: com.sun.star.style.ParagraphProperties
                    XPropertySet xSet = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, xInfo );
                    // Set the justification to be center justified
                    xSet.setPropertyValue ( "ParaAdjust", com.sun.star.style.ParagraphAdjust.CENTER );
                }
            }
            /* I tried to use the hacky, deprecated XRelativeTextContentInsert, but couldn't get it to work
            XTextContent xContent = (XTextContent) UnoRuntime.queryInterface ( XTextContent.class,
                           xFactory.createInstance ( "com.sun.star.text.Paragraph" ) );

            XRelativeTextContentInsert xRelative = (XRelativeTextContentInsert ) UnoRuntime.queryInterface ( XRelativeTextContentInsert.class, xDocText );

            XTextContent xEnd = (XTextContent) UnoRuntime.queryInterface ( XTextContent.class, xDocText.getEnd() );

            xRelative.insertTextContentAfter ( xContent, xDateField );
            XPropertySet xParaProp = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, myDoc );
            xParaProp.setPropertyValue ( "ParaAdjust", com.sun.star.style.ParagraphAdjust.LEFT );
            */
        }
        catch ( Exception e )
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "XParagraph example complete. Word count of document is now " + GetWordCount() );
    }
    /** This method shows how to create and manipulate text frames
     */
    protected void TextFrameExample ()
    {
        try
        {
            // Use the document's factory to create a new text frame and immediately access
            // it's XTextFrame interface
            XTextFrame xFrame = (XTextFrame) UnoRuntime.queryInterface ( XTextFrame.class,
                                     mxDocFactory.createInstance ( "com.sun.star.text.TextFrame" ) );
            // Access the XShape interface of the TextFrame
            XShape xShape = (XShape) UnoRuntime.queryInterface(XShape.class, xFrame);
            // Access the XPropertySet interface of the TextFrame
            XPropertySet xFrameProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xFrame );

            // Set the size of the new Text Frame using the XShape's 'setSize' method
            Size aSize = new Size();
            aSize.Height = 400;
            aSize.Width = 15000;
            xShape.setSize(aSize);
            // Set the AnchorType to com.sun.star.text.TextContentAnchorType.AS_CHARACTER
            xFrameProps.setPropertyValue( "AnchorType", TextContentAnchorType.AS_CHARACTER );
            // Go to the end of the text document
            mxDocCursor.gotoEnd( false );
            // Insert a new paragraph
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );
            // Then insert the new frame
            mxDocText.insertTextContent(mxDocCursor, xFrame, false);

            // Access the XText interface of the text contained within the frame
            XText xFrameText = xFrame.getText();
            // Create a TextCursor over the frame's contents
            XTextCursor xFrameCursor = xFrameText.createTextCursor();
            // Insert some text into the frame
            xFrameText.insertString( xFrameCursor, "The first line in the newly created text frame.", false );
            xFrameText.insertString( xFrameCursor, "\nThe second line in the new text frame.", false );
            // Insert a paragraph break into the document (not the frame)
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );
         }
        catch (Exception e)
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "XTextFrame example complete. Word count of document is now " + GetWordCount() );
    }
    /** This method returns a random double which isn't too high or too low
     */
    protected double getRandomDouble ()
    {
        return ( ( maRandom.nextInt() % 1000 ) * maRandom.nextDouble () );
    }
    /** This method shows how to create and insert a text table, as well as insert text and formulae
        into the cells of the table
     */
    protected void TextTableExample ()
    {
        try
        {
            // Create a new table from the document's factory
            XTextTable xTable = (XTextTable) UnoRuntime.queryInterface( XTextTable.class,
                mxDocFactory .createInstance( "com.sun.star.text.TextTable" ) );

            // Specify that we want the table to have 4 rows and 4 columns
            xTable.initialize( 4, 4 );

            // Insert the table into the document
            mxDocText.insertTextContent( mxDocCursor, xTable, false);
            // Get an XIndexAccess of the table rows
            XIndexAccess xRows = xTable.getRows();

            // Access the property set of the first row (properties listed in service description:
            // com.sun.star.text.TextTableRow)
            XPropertySet xRow = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, xRows.getByIndex ( 0 ) );
            // If BackTransparant is false, then the background color is visible
            xRow.setPropertyValue( "BackTransparent", new Boolean(false));
            // Specify the color of the background to be dark blue
            xRow.setPropertyValue( "BackColor", new Integer(6710932));

            // Access the property set of the whole table
            XPropertySet xTableProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xTable );
            // We want visible background colors
            xTableProps.setPropertyValue( "BackTransparent", new Boolean(false));
            // Set the background colour to light blue
            xTableProps.setPropertyValue( "BackColor", new Integer(13421823));

            // set the text (and text colour) of all the cells in the first row of the table
            insertIntoCell( "A1", "First Column", xTable );
            insertIntoCell( "B1", "Second Column", xTable );
            insertIntoCell( "C1", "Third Column", xTable );
            insertIntoCell( "D1", "Results", xTable );

            // Insert random numbers into the first this three cells of each
            // remaining row
            xTable.getCellByName( "A2" ).setValue( getRandomDouble() );
            xTable.getCellByName( "B2" ).setValue( getRandomDouble() );
            xTable.getCellByName( "C2" ).setValue( getRandomDouble() );

            xTable.getCellByName( "A3" ).setValue( getRandomDouble() );
            xTable.getCellByName( "B3" ).setValue( getRandomDouble() );
            xTable.getCellByName( "C3" ).setValue( getRandomDouble() );

            xTable.getCellByName( "A4" ).setValue( getRandomDouble() );
            xTable.getCellByName( "B4" ).setValue( getRandomDouble() );
            xTable.getCellByName( "C4" ).setValue( getRandomDouble() );

            // Set the last cell in each row to be a formula that calculates
            // the sum of the first three cells
            xTable.getCellByName( "D2" ).setFormula( "sum <A2:C2>" );
            xTable.getCellByName( "D3" ).setFormula( "sum <A3:C3>" );
            xTable.getCellByName( "D4" ).setFormula( "sum <A4:C4>" );
        }
        catch (Exception e)
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "XTextTable example complete. Word count of document is now " + GetWordCount() );
    }
    /** This method demonstrates how to create linked and unlinked sections
     */
    protected void TextSectionExample ()
    {
        try
        {
            // Go to the end of the document
            mxDocCursor.gotoEnd( false );
            // Insert two paragraph breaks
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, true );

            // Create a new TextSection from the document factory and access it's XNamed interface
            XNamed xChildNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class,
                mxDocFactory.createInstance( "com.sun.star.text.TextSection" ) );
            // Set the new sections name to 'Child_Section'
            xChildNamed.setName ( "Child_Section" );

            // Access the Child_Section's XTextContent interface and insert it into the document
            XTextContent xChildSection = (XTextContent) UnoRuntime.queryInterface( XTextContent.class, xChildNamed );
            mxDocText.insertTextContent ( mxDocCursor, xChildSection, false );

            // Access the XParagraphCursor interface of our text cursor
            XParagraphCursor xParaCursor = (XParagraphCursor)
                UnoRuntime.queryInterface( XParagraphCursor.class, mxDocCursor );

            // Go back one paragraph (into Child_Section)
            xParaCursor.gotoPreviousParagraph ( false );

            // Insert a string into the Child_Section
            mxDocText.insertString ( mxDocCursor, "This is a test", false );

            // Go to the end of the document
            mxDocCursor.gotoEnd( false );

            // Go back two paragraphs
            xParaCursor.gotoPreviousParagraph ( false );
            xParaCursor.gotoPreviousParagraph ( false );
            // Go to the end of the document, selecting the two paragraphs
            mxDocCursor.gotoEnd ( true );

            // Create another text section and access it's XNamed interface
            XNamed xParentNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class,
                mxDocFactory.createInstance( "com.sun.star.text.TextSection" ) );

            // Set this text section's name to Parent_Section
            xParentNamed.setName ( "Parent_Section" );

            // Access the Parent_Section's XTextContent interface ...
            XTextContent xParentSection = (XTextContent) UnoRuntime.queryInterface( XTextContent.class, xParentNamed );
            // ...and insert it into the document
            mxDocText.insertTextContent ( mxDocCursor, xParentSection, false );

            // Go to the end of the document
            mxDocCursor.gotoEnd ( false );
            // Insert a new paragraph
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );
            // And select the new pargraph
            xParaCursor.gotoPreviousParagraph ( true );

            // Create a new Text Section and access it's XNamed interface
            XNamed xLinkNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class,
                mxDocFactory.createInstance( "com.sun.star.text.TextSection" ) );
            // Set the new text section's name to Linked_Section
            xLinkNamed.setName ( "Linked_Section" );

            // Access the Linked_Section's XTextContent interface
            XTextContent xLinkedSection = (XTextContent) UnoRuntime.queryInterface( XTextContent.class, xLinkNamed );
            // And insert the Linked_Section into the document
            mxDocText.insertTextContent ( mxDocCursor, xLinkedSection, false );

            // Access the Linked_Section's XPropertySet interface
            XPropertySet xLinkProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xLinkNamed );
            // Set the linked section to be linked to the Child_Section
            xLinkProps.setPropertyValue ( "LinkRegion", "Child_Section" );

            // Access the XPropertySet interface of the Child_Section
            XPropertySet xChildProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, xChildNamed );
            // Set the Child_Section's background colour to blue
            xChildProps.setPropertyValue( "BackColor", new Integer(13421823));

            // Refresh the document, so the linked section matches the Child_Section
            XRefreshable xRefresh = (XRefreshable) UnoRuntime.queryInterface( XRefreshable.class, mxDoc );
            xRefresh.refresh();
        }
        catch (Exception e)
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "XTextSection example complete. Word count of document is now " + GetWordCount() );
    }
    /** This method demonstrates the XTextColumns interface and how to insert a blank paragraph
        using the XRelativeTextContentInsert interface
     */
    protected void TextColumnsExample ()
    {
        try
        {
            // Go to the end of the doucment
            mxDocCursor.gotoEnd( false );
            // insert a new paragraph
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );

            // insert the string 'I am a fish.' 100 times
            for ( int i = 0 ; i < 100 ; ++i )
            {
                mxDocText.insertString ( mxDocCursor, "I am a fish.", false );
            }
            // insert a paragraph break after the text
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );

            // Get the XParagraphCursor interface of our text cursor
            XParagraphCursor xParaCursor = (XParagraphCursor)
                UnoRuntime.queryInterface( XParagraphCursor.class, mxDocCursor );
            // Jump back before all the text we just inserted
            xParaCursor.gotoPreviousParagraph ( false );
            xParaCursor.gotoPreviousParagraph ( false );

            // Insert a string at the beginning of the block of text
            mxDocText.insertString ( mxDocCursor, "Fish section begins:", false );

            // Then select all of the text
            xParaCursor.gotoNextParagraph ( true );
            xParaCursor.gotoNextParagraph ( true );

            // Create a new text section and get it's XNamed interface
            XNamed xSectionNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class,
                mxDocFactory.createInstance( "com.sun.star.text.TextSection" ) );

            // Set the name of our new section (appropiately) to 'Fish'
            xSectionNamed.setName ( "Fish" );

            // Create the TextColumns service and get it's XTextColumns interface
            XTextColumns xColumns = (XTextColumns) UnoRuntime.queryInterface( XTextColumns.class,
                mxDocFactory.createInstance ( "com.sun.star.text.TextColumns" ) );

            // We want three columns
            xColumns.setColumnCount ( (short) 3 );

            // Get the TextColumns, and make the middle one narrow with a larger margin
            // on the left than the right
            TextColumn[]  aSequence = xColumns.getColumns ();
            aSequence[1].Width /= 2;
            aSequence[1].LeftMargin = 350;
            aSequence[1].RightMargin = 200;
            // Set the updated TextColumns back to the XTextColumns
            xColumns.setColumns ( aSequence );

            // Get the property set interface of our 'Fish' section
            XPropertySet xSectionProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, xSectionNamed );

            // Set the columns to the Text Section
            xSectionProps.setPropertyValue ( "TextColumns", xColumns );

            // Get the XTextContent interface of our 'Fish' section
            mxFishSection = (XTextContent) UnoRuntime.queryInterface( XTextContent.class, xSectionNamed );

            // Insert the 'Fish' section over the currently selected text
            mxDocText.insertTextContent ( mxDocCursor, mxFishSection, true );

            // Get the wonderful XRelativeTextContentInsert interface
            XRelativeTextContentInsert xRelative =
                (XRelativeTextContentInsert ) UnoRuntime.queryInterface ( XRelativeTextContentInsert.class, mxDocText );

            // Create a new empty paragraph and get it's XTextContent interface
            XTextContent xNewPara = (XTextContent) UnoRuntime.queryInterface( XTextContent.class,
                mxDocFactory.createInstance( "com.sun.star.text.Paragraph" ) );

            // Insert the empty paragraph after the fish Text Section
            xRelative.insertTextContentAfter ( xNewPara, mxFishSection );
        }
        catch (Exception e)
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "XTextColumns example complete. Word count of document is now " + GetWordCount() );
    }
    /** This method demonstrates how to set numbering types and numbering levels using the
        com.sun.star.text.NumberingRules service
     */
    protected void NumberingExample ()
    {
        try
        {
            // Go to the end of the document
            mxDocCursor.gotoEnd( false );
            // Get the RelativeTextContentInsert interface of the document
            XRelativeTextContentInsert xRelative =
                (XRelativeTextContentInsert ) UnoRuntime.queryInterface ( XRelativeTextContentInsert.class, mxDocText );

            // Use the document's factory to create the NumberingRules service, and get it's XIndexAccess interface
            XIndexAccess xNum = (XIndexAccess) UnoRuntime.queryInterface( XIndexAccess.class,
                mxDocFactory.createInstance( "com.sun.star.text.NumberingRules" ) );

            // Also get the NumberingRule's XIndexReplace interface
            XIndexReplace xReplace = (XIndexReplace) UnoRuntime.queryInterface( XIndexReplace.class, xNum );

            // Create an array of XPropertySets, one for each of the three paragraphs we're about to create
            XPropertySet xParas[] = new XPropertySet [ 3 ];
            for ( int i = 0 ; i < 3 ; ++ i )
            {
                // Create a new paragraph
                XTextContent xNewPara = (XTextContent) UnoRuntime.queryInterface( XTextContent.class,
                    mxDocFactory.createInstance( "com.sun.star.text.Paragraph" ) );

                // Get the XPropertySet interface of the new paragraph and put it in our array
                xParas[i] = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, xNewPara );

                // Insert the new paragraph into the document after the fish section. As it is an insert
                // relative to the fish section, the first paragraph inserted will be below the next two
                xRelative.insertTextContentAfter ( xNewPara, mxFishSection );

                // Seperate from the above, but also needs to be done three times

                // Get the PropertyValue sequence for this numbering level
                PropertyValue [] aProps = (PropertyValue [] ) xNum.getByIndex ( i );

                // Iterate over the PropertyValue's for this numbering level, looking for the
                // 'NumberingType' property
                for ( int j = 0 ; j < aProps.length ; ++j )
                {
                    if ( aProps[j].Name.equals ( "NumberingType" ) )
                    {
                        // Once we find it, set it's value to a new type, dependant on which
                        // numbering level we're currently on
                        switch ( i )
                        {
                            case 0 : aProps[j].Value = new Short ( NumberingType.ROMAN_UPPER );
                            break;
                            case 1 : aProps[j].Value = new Short ( NumberingType.CHARS_UPPER_LETTER );
                            break;
                            case 2 : aProps[j].Value = new Short ( NumberingType.ARABIC );
                            break;
                        }
                        // Put the updated PropertyValue sequence back into the NumberingRules service
                        xReplace.replaceByIndex ( i, aProps );
                        break;
                    }
                }
            }
            // Get the XParagraphCursor interface of our text cursro
            XParagraphCursor xParaCursor = (XParagraphCursor)
                UnoRuntime.queryInterface( XParagraphCursor.class, mxDocCursor );
            // Go to the end of the document, then select the preceding paragraphs
            mxDocCursor.gotoEnd ( false );
            xParaCursor.gotoPreviousParagraph ( false );
            xParaCursor.gotoPreviousParagraph ( true );
            xParaCursor.gotoPreviousParagraph ( true );

            // Get the XPropertySet of the cursor's currently selected text
            XPropertySet xCursorProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, mxDocCursor );

            // Set the updated Numbering rules to the cursor's property set
            xCursorProps.setPropertyValue ( "NumberingRules", xNum );
            mxDocCursor.gotoEnd( false );

            // Set the first paragraph that was inserted to a numbering level of 2 (thus it will
            // have Arabic style numbering)
            xParas[0].setPropertyValue ( "NumberingLevel", new Short ( (short) 2 ) );

            // Set the second paragraph that was inserted to a numbering level of 1 (thus it will
            // have 'Chars Upper Letter' style numbering)
            xParas[1].setPropertyValue ( "NumberingLevel", new Short ( (short) 1 ) );

            // Set the third paragraph that was inserted to a numbering level of 0 (thus it will
            // have 'Chars Upper Letter' style numbering)
            xParas[2].setPropertyValue ( "NumberingLevel", new Short ( (short) 0 ) );
        }
        catch (Exception e)
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "Numbering rules example complete. Word count of document is now " + GetWordCount() );
    }

    /** This method demonstrates how to create, insert and apply new styles
     */
    protected void StylesExample ()
    {
        try
        {
            // Go to the end of the document
            mxDocCursor.gotoEnd( false );

            // Insert two paragraph breaks
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );

            // Create a new style from the document's factory
            XStyle xStyle = (XStyle) UnoRuntime.queryInterface( XStyle.class,
                    mxDocFactory.createInstance( "com.sun.star.style.ParagraphStyle" ) );

            // Access the XPropertySet interface of the new style
            XPropertySet xStyleProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, xStyle );

            // Give the new style a light blue background
            xStyleProps.setPropertyValue ( "ParaBackColor", new Integer ( 13421823 ) );

            // Get the StyleFamiliesSupplier interface of the document
            XStyleFamiliesSupplier xSupplier = ( XStyleFamiliesSupplier ) UnoRuntime.queryInterface
                            ( XStyleFamiliesSupplier.class, mxDoc );

            // Use the StyleFamiliesSupplier interface to get the XNameAccess interface of the
            // actual style families
            XNameAccess xFamilies = ( XNameAccess ) UnoRuntime.queryInterface ( XNameAccess.class,
                            xSupplier.getStyleFamilies() );

            // Access the 'ParagraphStyles' Family
            XNameContainer xFamily = (XNameContainer ) UnoRuntime.queryInterface ( XNameContainer.class,
                            xFamilies.getByName ( "ParagraphStyles" ) );

            // Insert the newly created style into the ParagraphStyles family
            xFamily.insertByName ( "All-Singing All-Dancing Style", xStyle );

            // Get the XParagraphCursor interface of the document cursor
            XParagraphCursor xParaCursor = (XParagraphCursor)
                UnoRuntime.queryInterface( XParagraphCursor.class, mxDocCursor );

            // Select the first paragraph inserted
            xParaCursor.gotoPreviousParagraph ( false );
            xParaCursor.gotoPreviousParagraph ( true );

            // Access the property set of the cursor selection
            XPropertySet xCursorProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, mxDocCursor );

            // Set the style of the cursor selection to our newly created style
            xCursorProps.setPropertyValue ( "ParaStyleName", "All-Singing All-Dancing Style" );

            // Go back to the end
            mxDocCursor.gotoEnd ( false );

            // Select the last paragraph in the document
            xParaCursor.gotoNextParagraph ( true );

            // And reset it's style to 'Standard' (the programmatic name for the default style)
            xCursorProps.setPropertyValue ( "ParaStyleName", "Standard" );

        }
        catch (Exception e)
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "XStyles example complete. Word count of document is now " + GetWordCount() );
    }

    /** This method demonstrates how to insert indices and index marks
     */
    protected void IndexExample ()
    {
        try
        {
            // Go to the end of the document
            mxDocCursor.gotoEnd( false );
            // Insert a new paragraph and position the cursor in it
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );
            XParagraphCursor xParaCursor = (XParagraphCursor)
                UnoRuntime.queryInterface( XParagraphCursor.class, mxDocCursor );
            xParaCursor.gotoPreviousParagraph ( false );

            // Create a new ContentIndexMark and get it's XPropertySet interface
            XPropertySet xEntry = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class,
                mxDocFactory.createInstance ( "com.sun.star.text.ContentIndexMark" ) );

            // Set the text to be displayed in the index
            xEntry.setPropertyValue ( "AlternativeText", "Big dogs! Falling on my head!" );

            // The Level property _must_ be set
            xEntry.setPropertyValue ( "Level", new Short ( (short) 1 ) );

            // Create a ContentIndex and access it's XPropertySet interface
            XPropertySet xIndex = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class,
                mxDocFactory.createInstance ( "com.sun.star.text.ContentIndex" ) );

            // Again, the Level property _must_ be set
            xIndex.setPropertyValue ( "Level", new Short ( (short) 10 ) );

            // Access the XTextContent interfaces of both the Index and the IndexMark
            XTextContent xIndexContent = (XTextContent) UnoRuntime.queryInterface(
                XTextContent.class, xIndex );
            XTextContent xEntryContent = (XTextContent) UnoRuntime.queryInterface(
                XTextContent.class, xEntry );

            // Insert both in the document
            mxDocText.insertTextContent ( mxDocCursor, xEntryContent, false );
            mxDocText.insertTextContent ( mxDocCursor, xIndexContent, false );

            // Get the XDocumentIndex interface of the Index
            XDocumentIndex xDocIndex = (XDocumentIndex) UnoRuntime.queryInterface(
                XDocumentIndex.class, xIndex );

            // And call it's update method
            xDocIndex.update();
        }
        catch (Exception e)
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "Index example complete. Word count of document is now " + GetWordCount() );
    }

    /** This method demonstrates how to create and insert reference marks, and GetReference Text Fields
     */
    protected void ReferenceExample ()
    {
        try
        {
            // Go to the end of the document
            mxDocCursor.gotoEnd( false );

            // Insert a paragraph break
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );

            // Get the Paragraph cursor
            XParagraphCursor xParaCursor = (XParagraphCursor)
                UnoRuntime.queryInterface( XParagraphCursor.class, mxDocCursor );

            // Move the cursor into the new paragraph
            xParaCursor.gotoPreviousParagraph ( false );

            // Create a new ReferenceMark and get it's XNamed interface
            XNamed xRefMark = (XNamed) UnoRuntime.queryInterface( XNamed.class,
                mxDocFactory.createInstance ( "com.sun.star.text.ReferenceMark" ) );

            // Set the name to TableHeader
            xRefMark.setName ( "TableHeader" );

            // Get the TextTablesSupplier interface of the document
            XTextTablesSupplier xTableSupplier = ( XTextTablesSupplier ) UnoRuntime.queryInterface
                            ( XTextTablesSupplier.class, mxDoc );

            // Get an XIndexAccess of TextTables
            XIndexAccess xTables = ( XIndexAccess ) UnoRuntime.queryInterface ( XIndexAccess.class,
                            xTableSupplier.getTextTables() );

            // We've only inserted one table, so get the first one from index zero
            XTextTable xTable = ( XTextTable ) UnoRuntime.queryInterface ( XTextTable.class,
                            xTables.getByIndex( 0 ) );

            // Get the first cell from the table
            XText xTableText = (XText) UnoRuntime.queryInterface(XText.class, xTable.getCellByName ( "A1" ) );

            // Get a text cursor for the first cell
            XTextCursor xTableCursor = xTableText.createTextCursor();

            // Get the XTextContent interface of the reference mark so we can insert it
            XTextContent xContent = ( XTextContent ) UnoRuntime.queryInterface ( XTextContent.class, xRefMark );

            // Insert the reference mark into the first cell of the table
            xTableText.insertTextContent ( xTableCursor, xContent, false );

            // Create a 'GetReference' text field to refer to the reference mark we just inserted,
            // and get it's XPropertySet interface
            XPropertySet xFieldProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class,
                mxDocFactory.createInstance ( "com.sun.star.text.TextField.GetReference" ) );

            // Get the XReferenceMarksSupplier interface of the document
            XReferenceMarksSupplier xRefSupplier = ( XReferenceMarksSupplier ) UnoRuntime.queryInterface
                            ( XReferenceMarksSupplier.class, mxDoc );

            // Get an XNameAccess which refers to all inserted reference marks
            XNameAccess xMarks = ( XNameAccess ) UnoRuntime.queryInterface ( XNameAccess.class,
                            xRefSupplier.getReferenceMarks() );

            // Put the names of each reference mark into an array of strings
            String[] aNames = xMarks.getElementNames();

            // Make sure that at least 1 reference mark actually exists (well, we just inserted one!)
            if ( aNames.length > 0 )
            {
                // Output the name of the first reference mark ('TableHeader')
                System.out.println ( "GetReference text field inserted for ReferenceMark : " + aNames[0] );

                // Set the SourceName of the GetReference text field to 'TableHeader'
                xFieldProps.setPropertyValue ( "SourceName", aNames[0] );

                // specify that the source is a reference mark (could also be a footnote, bookmark or sequence field )
                xFieldProps.setPropertyValue ( "ReferenceFieldSource", new Short ( ReferenceFieldSource.REFERENCE_MARK ) );

                // We want the reference displayed as 'above' or 'below'
                xFieldProps.setPropertyValue ( "ReferenceFieldPart", new Short ( ReferenceFieldPart.UP_DOWN ) );

                // Get the XTextContent interface of the GetReference text field
                XTextContent xRefContent = (XTextContent) UnoRuntime.queryInterface( XTextContent.class, xFieldProps );

                // Go to the end of the document
                mxDocCursor.gotoEnd( false );

                // Make some text to precede the reference
                mxDocText.insertString ( mxDocText.getEnd(), "The table ", false );

                // Insert the text field
                mxDocText.insertTextContent ( mxDocText.getEnd(), xRefContent, false );

                // And some text after the reference..
                mxDocText.insertString ( mxDocText.getEnd(), " contains the sum of some random numbers.", false );

                // Refresh the document
                XRefreshable xRefresh = (XRefreshable) UnoRuntime.queryInterface( XRefreshable.class, mxDoc );
                xRefresh.refresh();
            }
        }
        catch (Exception e)
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "Reference example complete. Word count of document is now " + GetWordCount() );
    }

    /** This method demonstrates how to create and insert footnotes, and how to access the
        XFootnotesSupplier interface of the document
     */
    protected void FootnoteExample ()
    {
        try
        {
            // Create a new footnote from the document factory and get it's
            // XFootnote interface
            XFootnote xFootnote = (XFootnote) UnoRuntime.queryInterface( XFootnote.class,
                mxDocFactory.createInstance ( "com.sun.star.text.Footnote" ) );

            // Set the label to 'Numbers'
            xFootnote.setLabel ( "Numbers" );

            // Get the footnotes XTextContent interface so we can...
            XTextContent xContent = ( XTextContent ) UnoRuntime.queryInterface ( XTextContent.class, xFootnote );

            // ...insert it into the document
            mxDocText.insertTextContent ( mxDocCursor, xContent, false );

            // Get the XFootnotesSupplier interface of the document
            XFootnotesSupplier xFootnoteSupplier = ( XFootnotesSupplier ) UnoRuntime.queryInterface
                            ( XFootnotesSupplier.class, mxDoc );

            // Get an XIndexAccess interface to all footnotes
            XIndexAccess xFootnotes = ( XIndexAccess ) UnoRuntime.queryInterface ( XIndexAccess.class,
                            xFootnoteSupplier.getFootnotes() );

            // Get the XFootnote interface to the first footnote inserted ('Numbers')
            XFootnote xNumbers = ( XFootnote ) UnoRuntime.queryInterface ( XFootnote.class,
                            xFootnotes.getByIndex( 0 ) );

            // Get the XSimpleText interface to the Footnote
            XSimpleText xSimple = (XSimpleText ) UnoRuntime.queryInterface ( XSimpleText.class, xNumbers );

            // Create a text cursor for the foot note text
            XTextRange xRange = (XTextRange ) UnoRuntime.queryInterface ( XTextRange.class, xSimple.createTextCursor() );

            // And insert the actual text of the footnote.
            xSimple.insertString ( xRange, "  The numbers were generated by using java.util.Random", false );
        }
        catch (Exception e)
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "Footnote example complete. Word count of document is now " + GetWordCount() );
    }


    /** This method demonstrates how to create and manipulate shapes, and how to access the draw page
        of the document to insert shapes
     */
    protected void DrawPageExample ()
    {
        try
        {
            // Go to the end of the document
            mxDocCursor.gotoEnd( false );

            // Insert two new paragraphs
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );

            // Get the XParagraphCursor interface of our document cursor
            XParagraphCursor xParaCursor = (XParagraphCursor)
                UnoRuntime.queryInterface( XParagraphCursor.class, mxDocCursor );

            // Position the cursor in the 2nd paragraph
            xParaCursor.gotoPreviousParagraph ( false );

            // Create a RectangleShape using the document factory
            XShape xRect = (XShape) UnoRuntime.queryInterface( XShape.class,
                mxDocFactory.createInstance ( "com.sun.star.drawing.RectangleShape" ) );

            // Create an EllipseShape using the document factory
            XShape xEllipse = (XShape) UnoRuntime.queryInterface( XShape.class,
                mxDocFactory.createInstance ( "com.sun.star.drawing.EllipseShape" ) );

            // Set the size of both the ellipse and the rectangle
            Size aSize = new Size();
            aSize.Height = 4000;
            aSize.Width = 10000;
            xRect.setSize(aSize);
            aSize.Height = 3000;
            aSize.Width = 6000;
            xEllipse.setSize ( aSize );

            // Set the position of the Rectangle to the right of the ellipse
            Point aPoint = new Point();
            aPoint.X = 6100;
            aPoint.Y = 0;
            xRect.setPosition ( aPoint );

            // Get the XPropertySet interfaces of both shapes
            XPropertySet xRectProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, xRect );
            XPropertySet xEllipseProps = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, xEllipse );

            // And set the AnchorTypes of both shapes to 'AT_PARAGRAPH'
            xRectProps.setPropertyValue ( "AnchorType", TextContentAnchorType.AT_PARAGRAPH );
            xEllipseProps.setPropertyValue ( "AnchorType", TextContentAnchorType.AT_PARAGRAPH );

            // Access the XDrawPageSupplier interface of the document
            XDrawPageSupplier xDrawPageSupplier = ( XDrawPageSupplier ) UnoRuntime.queryInterface
                            ( XDrawPageSupplier.class, mxDoc );

            // Get the XShapes interface of the draw page
            XShapes xShapes = ( XShapes ) UnoRuntime.queryInterface ( XShapes.class, xDrawPageSupplier.getDrawPage () );

            // Add both shapes
            xShapes.add ( xEllipse );
            xShapes.add ( xRect );

            /*

            This doesn't work, I am assured that FME and AMA are fixing it.

            XShapes xGrouper = (XShapes) UnoRuntime.queryInterface( XShapes.class,
                mxDocFactory.createInstance ( "com.sun.star.drawing.GroupShape" ) );

            XShape xGrouperShape = (XShape) UnoRuntime.queryInterface( XShape.class, xGrouper );
            xShapes.add ( xGrouperShape );

            xGrouper.add ( xRect );
            xGrouper.add ( xEllipse );

            XShapeGrouper xShapeGrouper = ( XShapeGrouper ) UnoRuntime.queryInterface
                ( XShapeGrouper.class, xShapes );
            xShapeGrouper.group ( xGrouper );
            */

        }
        catch (Exception e)
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "Drawpage example complete. Word count of document is now " + GetWordCount() );
    }

    /** This example demonstrates the use of the AutoTextContainer, AutoTextGroup and AutoTextEntry services
        and shows how to create, insert and modify auto text blocks
     */
    protected void AutoTextExample ()
    {
        try
        {
            // Go to the end of the document
            mxDocCursor.gotoEnd( false );

            // Insert two paragraphs
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );

            // Position the cursor in the second paragraph
            XParagraphCursor xParaCursor = (XParagraphCursor)
                UnoRuntime.queryInterface( XParagraphCursor.class, mxDocCursor );
            xParaCursor.gotoPreviousParagraph ( false );

            // Get an XNameAccess interface to all auto text groups from the document factory
            XNameAccess xContainer = (XNameAccess) UnoRuntime.queryInterface( XNameAccess.class,
                mxFactory.createInstance ( "com.sun.star.text.AutoTextContainer" ) );

            // Create a new table at the document factory
            XTextTable xTable = (XTextTable) UnoRuntime.queryInterface( XTextTable.class,
                mxDocFactory .createInstance( "com.sun.star.text.TextTable" ) );

            // Store the names of all auto text groups in an array of strings
            String[] aGroupNames = xContainer.getElementNames();

            // Make sure we have at least one group name
            if ( aGroupNames.length > 0 )
            {
                // initialise the table to have a row for every autotext group in a single column + one
                // additional row for a header
                xTable.initialize( aGroupNames.length+1,1);

                // Access the XPropertySet of the table
                XPropertySet xTableProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xTable );

                // We want a visible background
                xTableProps.setPropertyValue( "BackTransparent", new Boolean(false));

                // We want the background to be light blue
                xTableProps.setPropertyValue( "BackColor", new Integer(13421823));

                // Inser the table into the document
                mxDocText.insertTextContent( mxDocCursor, xTable, false);

                // Get an XIndexAccess to all table rows
                XIndexAccess xRows = xTable.getRows();

                // Get the first row in the table
                XPropertySet xRow = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, xRows.getByIndex ( 0 ) );

                // We want the background of the first row to be visible too
                xRow.setPropertyValue( "BackTransparent", new Boolean(false));

                // And let's make it dark blue
                xRow.setPropertyValue( "BackColor", new Integer(6710932));

                // Put a description of the table contents into the first cell
                insertIntoCell( "A1", "AutoText Groups", xTable);

                // Create a table cursor pointing at the second cell in the first column
                XTextTableCursor xTableCursor = xTable.createCursorByCellName ( "A2" );

                // Loop over the group names
                for ( int i = 0 ; i < aGroupNames.length ; i ++ )
                {
                    // Get the name of the current cell
                    String sCellName = xTableCursor.getRangeName ();

                    // Get the XText interface of the current cell
                    XText xCellText = (XText) UnoRuntime.queryInterface ( XText.class,
                            xTable.getCellByName ( sCellName ) );

                    // Set the cell contents of the current cell to be the name of the of an
                    // autotext group
                    xCellText.setString ( aGroupNames[i] );

                    // Access the autotext gruop with this name
                    XAutoTextGroup xGroup = ( XAutoTextGroup ) UnoRuntime.queryInterface ( XAutoTextGroup.class,
                            xContainer.getByName ( aGroupNames[i] ) );

                    // Get the titles of each autotext block in this group
                    String [] aBlockNames = xGroup.getTitles();

                    // Make sure that the autotext group contains at least one block
                    if ( aBlockNames.length > 0 )
                    {
                        // Split the current cell vertically into two seperate cells
                        xTableCursor.splitRange ( (short) 1, false );

                        // Put the cursor in the newly created right hand cell and select it
                        xTableCursor.goRight ( (short) 1, false );

                        // Split this cell horizontally to make a seperate cell for each
                        // Autotext block
                        if ( ( aBlockNames.length -1 ) > 0 )
                            xTableCursor.splitRange ( (short) (aBlockNames.length - 1), true );

                        // loop over the block names
                        for ( int j = 0 ; j < aBlockNames.length ; j ++ )
                        {
                            // Get the XText interface of the current cell
                            xCellText = (XText) UnoRuntime.queryInterface ( XText.class,
                                    xTable.getCellByName ( xTableCursor.getRangeName() ) );

                            // Set the text contents of the current cell to the title of an
                            // Autotext block
                            xCellText.setString ( aBlockNames[j] );

                            // Move the cursor down one cell
                            xTableCursor.goDown( (short)1, false);
                        }
                    }
                    // Go back to the cell we originally split
                    xTableCursor.gotoCellByName ( sCellName, false );

                    // Go down one cell
                    xTableCursor.goDown( (short)1, false);
                }

                XAutoTextGroup xGroup;
                String [] aBlockNames;

                // Add a depth so that we only generate 200 numbers before giving up on finding
                // a random autotext group that contains autotext blocks
                int nDepth = 0;
                do
                {
                    // Generate a random, positive number which is lower than the number of autotext groups
                    int nRandom = Math.abs ( maRandom.nextInt() % aGroupNames.length );

                    // Get the autotext group at this name
                    xGroup 	= ( XAutoTextGroup ) UnoRuntime.queryInterface ( XAutoTextGroup.class,
                        xContainer.getByName ( aGroupNames[ nRandom ] ) );

                    // Fill our string array with the names of all the blocks in this group
                    aBlockNames = xGroup.getElementNames();

                    // increment our depth counter
                    ++nDepth;
                }
                while ( nDepth < 200 && aBlockNames.length == 0 );

                // If we managed to find a group containg blocks...
                if ( aBlockNames.length > 0 )
                {
                    // Pick a random block in this group and get it's XAutoTextEntry interface
                    int nRandom = Math.abs ( maRandom.nextInt() % aBlockNames.length );
                    XAutoTextEntry xEntry = ( XAutoTextEntry ) UnoRuntime.queryInterface ( XAutoTextEntry.class,
                            xGroup.getByName ( aBlockNames[ nRandom ] ) );

                    // insert the modified autotext block at the end of the document
                    xEntry.applyTo ( mxDocCursor );

                    // Get the titles of all text blocks in this AutoText group
                    String [] aBlockTitles = xGroup.getTitles();

                    // Get the XNamed interface of the autotext group
                    XNamed xGroupNamed = ( XNamed ) UnoRuntime.queryInterface ( XNamed.class, xGroup );

                    // Output the short cut and title of the random block and the name of the group it's from
                    System.out.println ( "Inserted the Autotext '" + aBlockTitles[nRandom] + "', shortcut '" + aBlockNames[nRandom] +
                                "' from group '" + xGroupNamed.getName() );
                }
            }

            // Go to the end of the document
            mxDocCursor.gotoEnd( false );
            // Insert new paragraph
            mxDocText.insertControlCharacter ( mxDocCursor, ControlCharacter.PARAGRAPH_BREAK, false );

            // Position cursor in new paragraph
            xParaCursor.gotoPreviousParagraph ( false );

            // Insert a string in the new paragraph
            mxDocText.insertString ( mxDocCursor, "Some text for a new autotext block", false );

            // Go to the end of the document
            mxDocCursor.gotoEnd( false );

            // Select the last paragraph in the document
            xParaCursor.gotoPreviousParagraph ( true );

            // Get the XAutoTextContainer interface of the AutoTextContainer service
            XAutoTextContainer xAutoTextCont = (XAutoTextContainer) UnoRuntime.queryInterface( XAutoTextContainer.class,
                xContainer );

            // If the APIExampleGroup already exists, remove it so we can add a new one
            if ( xContainer.hasByName ( "APIExampleGroup" ) )
                xAutoTextCont.removeByName ( "APIExampleGroup" );

            // Create a new auto-text group called APIExampleGroup
            XAutoTextGroup xNewGroup = xAutoTextCont.insertNewByName ( "APIExampleGroup" );

            // Create and insert a new auto text entry containing the current cursor selection
            XAutoTextEntry xNewEntry = xNewGroup.insertNewByName ( "NAE", "New AutoTextEntry", xParaCursor );

            // Get the XSimpleText and XText interfaces of the new autotext block
            XSimpleText xSimpleText = ( XSimpleText ) UnoRuntime.queryInterface ( XSimpleText.class, xNewEntry );
            XText xText = ( XText ) UnoRuntime.queryInterface ( XText.class, xNewEntry );

            // Insert a string at the beginning of the autotext block
            xSimpleText.insertString ( xText.getStart(), "This string was inserted using the API!\n\n", false );

            // Note, the new modified autotext isn't written to disk until the destructor of the SwXAutoTextEntry
            // inside the writer is called, therefore when this example has finished executing, the file on disk
            // will correctly contain both 'This string was inserted using the API!\n\nSome text for a new autotext block'
            // but as there is now way in java to actually call the destructor, we can't be sure when the garbage
            // collector will delete the object and thus write the modifications to disk.
            //
            // This problem indicates that XAutoTextEntry needs a new 'save' or 'commitChanges' method, but we don't
            // have one just yet...
        }
        catch (Exception e)
        {
            e.printStackTrace ( System.out );
        }
        System.out.println ( "AutoText example complete. Word count of document is now " + GetWordCount() );
    }

    /** Simple method to query the text document's property set for the current word count
     */
    protected long GetWordCount()
    {
        long nCount = 0;
        try
        {
            // Query for the property
            Object aCountObj = mxDocProps.getPropertyValue ( "WordCount" );
            // Use the AnyConverter to convert it into a long
            nCount = AnyConverter.toLong ( aCountObj );
        }
        catch ( Exception e )
        {
            e.printStackTrace ( System.out );
        }
        return nCount;
    }

    /** This method sets the text colour of the cell refered to by sCellName to white and inserts
        the string sText in it
     */
    public static void insertIntoCell(String sCellName, String sText, XTextTable xTable)
    {
        // Access the XText interface of the cell referred to by sCellName
        XText xCellText = (XText) UnoRuntime.queryInterface(XText.class, xTable.getCellByName ( sCellName ) );
        // Get a text cursor referring to the cell
        XTextCursor xCellCursor = xCellText.createTextCursor();
        // Get the property set of the cell's TextCursor
        XPropertySet xCellCursorProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xCellCursor );

        try
        {
            // Set the colour of the text to white
            xCellCursorProps.setPropertyValue( "CharColor", new Integer ( 16777215 ) );
        }
        catch ( Exception e)
        {
            e.printStackTrace ( System.out );
        }
        // Set the text in the cell to sText
        xCellText.setString( sText );
    }
}


