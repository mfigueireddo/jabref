package org.jabref.gui.openoffice;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Supplier;
import java.util.stream.Collectors;

import org.jabref.gui.DialogService;
import org.jabref.gui.StateManager;
import org.jabref.logic.JabRefException;
import org.jabref.logic.citationstyle.CitationStyle;
import org.jabref.logic.l10n.Localization;
import org.jabref.logic.openoffice.NoDocumentFoundException;
import org.jabref.logic.openoffice.OpenOfficePreferences;
import org.jabref.logic.openoffice.action.EditInsert;
import org.jabref.logic.openoffice.action.EditMerge;
import org.jabref.logic.openoffice.action.EditSeparate;
import org.jabref.logic.openoffice.action.ExportCited;
import org.jabref.logic.openoffice.action.ManageCitations;
import org.jabref.logic.openoffice.action.Update;
import org.jabref.logic.openoffice.frontend.OOFrontend;
import org.jabref.logic.openoffice.frontend.RangeForOverlapCheck;
import org.jabref.logic.openoffice.oocsltext.CSLCitationOOAdapter;
import org.jabref.logic.openoffice.oocsltext.CSLUpdateBibliography;
import org.jabref.logic.openoffice.style.JStyle;
import org.jabref.logic.openoffice.style.OOStyle;
import org.jabref.model.database.BibDatabase;
import org.jabref.model.database.BibDatabaseContext;
import org.jabref.model.entry.BibEntry;
import org.jabref.model.entry.BibEntryTypesManager;
import org.jabref.model.openoffice.CitationEntry;
import org.jabref.model.openoffice.rangesort.FunctionalTextViewCursor;
import org.jabref.model.openoffice.style.CitationGroupId;
import org.jabref.model.openoffice.style.CitationType;
import org.jabref.model.openoffice.uno.CreationException;
import org.jabref.model.openoffice.uno.NoDocumentException;
import org.jabref.model.openoffice.uno.UnoCrossRef;
import org.jabref.model.openoffice.uno.UnoCursor;
import org.jabref.model.openoffice.uno.UnoRedlines;
import org.jabref.model.openoffice.uno.UnoStyle;
import org.jabref.model.openoffice.uno.UnoUndo;
import org.jabref.model.openoffice.util.OOResult;
import org.jabref.model.openoffice.util.OOVoidResult;

import com.airhacks.afterburner.injection.Injector;
import com.sun.star.beans.IllegalTypeException;
import com.sun.star.beans.NotRemoveableException;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/// Class for manipulating the Bibliography of the currently started document in OpenOffice.
public class OOBibBase {

    private static final Logger LOGGER = LoggerFactory.getLogger(OOBibBase.class);

    private final DialogService dialogService;

    private final OOBibBaseConnect connection;

    private final OpenOfficePreferences openOfficePreferences;

    private CSLCitationOOAdapter cslCitationOOAdapter;
    private CSLUpdateBibliography cslUpdateBibliography;

    private final OOBibBaseGUI gui;

    public OOBibBase(Path loPath, DialogService dialogService, OpenOfficePreferences openOfficePreferences)
            throws
            BootstrapException,
            CreationException, IOException, InterruptedException {

        this.dialogService = dialogService;
        this.connection = new OOBibBaseConnect(loPath, dialogService);
        this.openOfficePreferences = openOfficePreferences;
        this.gui = new OOBibBaseGUI(this, dialogService);
    }

    public OOBibBaseGUI getGUI(){
        return this.gui;
    }

    private void initializeCitationAdapter(XTextDocument doc) throws WrappedTargetException, NoSuchElementException {
        if (cslCitationOOAdapter == null) {
            StateManager stateManager = Injector.instantiateModelOrService(StateManager.class);
            Supplier<List<BibDatabaseContext>> databasesSupplier = stateManager::getOpenDatabases;
            cslCitationOOAdapter = new CSLCitationOOAdapter(doc, databasesSupplier, openOfficePreferences, Injector.instantiateModelOrService(BibEntryTypesManager.class));
            cslUpdateBibliography = new CSLUpdateBibliography();
        }
    }

    public void guiActionSelectDocument(boolean autoSelectForSingle) throws WrappedTargetException, NoSuchElementException {
        final String errorTitle = Localization.lang("Problem connecting");

        try {
            connection.selectDocument(autoSelectForSingle);
        } catch (NoDocumentFoundException ex) {
            OOError.from(ex).showErrorDialog(dialogService);
        } catch (DisposedException ex) {
            OOError.from(ex).setTitle(errorTitle).showErrorDialog(dialogService);
        } catch (WrappedTargetException
                 | IndexOutOfBoundsException
                 | NoSuchElementException ex) {
            LOGGER.warn(errorTitle, ex);
            OOError.fromMisc(ex).setTitle(errorTitle).showErrorDialog(dialogService);
        }

        if (isConnectedToDocument()) {
            initializeCitationAdapter(this.getXTextDocument().get());
            dialogService.notify(Localization.lang("Connected to document") + ": "
                    + this.getCurrentDocumentTitle().orElse(""));
        }
    }

    /// A simple test for document availability.
    ///
    /// See also `isDocumentConnectionMissing` for a test actually attempting to use the connection.
    public boolean isConnectedToDocument() {
        return this.connection.isConnectedToDocument();
    }

    /// @return true if we are connected to a document
    public boolean isDocumentConnectionMissing() {
        return this.connection.isDocumentConnectionMissing();
    }

    /// Either return an XTextDocument or return JabRefException.
    public OOResult<XTextDocument, OOError> getXTextDocument() {
        return this.connection.getXTextDocument();
    }

    /// The title of the current document, or Optional.empty()
    public Optional<String> getCurrentDocumentTitle() {
        return this.connection.getCurrentDocumentTitle();
    }

    /* ******************************************************
     *
     *  Tools to collect and show precondition test results
     *
     * ******************************************************/

    OOVoidResult<OOError> collectResults(String errorTitle, List<OOVoidResult<OOError>> results) {
        String msg = results.stream()
                            .filter(OOVoidResult::isError)
                            .map(e -> e.getError().getLocalizedMessage())
                            .collect(Collectors.joining("\n\n"));
        if (msg.isEmpty()) {
            return OOVoidResult.ok();
        } else {
            return OOVoidResult.error(new OOError(errorTitle, msg));
        }
    }

    /// Get the cursor positioned by the user for inserting text.
    OOResult<XTextCursor, OOError> getUserCursorForTextInsertion(XTextDocument doc, String errorTitle) {
        // Get the cursor positioned by the user.
        XTextCursor cursor = UnoCursor.getViewCursor(doc).orElse(null);

        // Check for crippled XTextViewCursor
        Objects.requireNonNull(cursor);
        try {
            cursor.getStart();
        } catch (com.sun.star.uno.RuntimeException ex) {
            String msg =
                    Localization.lang("Please move the cursor"
                            + " to the location for the new citation.") + "\n"
                            + Localization.lang("I cannot insert to the cursor's current location.");
            return OOResult.error(new OOError(errorTitle, msg, ex));
        }
        return OOResult.ok(cursor);
    }

    /// This may move the view cursor.
    OOResult<FunctionalTextViewCursor, OOError> getFunctionalTextViewCursor(XTextDocument doc, String errorTitle) {
        String messageOnFailureToObtain =
                Localization.lang("Please move the cursor into the document text.")
                        + "\n"
                        + Localization.lang("To get the visual positions of your citations"
                        + " I need to move the cursor around,"
                        + " but could not get it.");
        OOResult<FunctionalTextViewCursor, String> result = FunctionalTextViewCursor.get(doc);
        if (result.isError()) {
            LOGGER.warn(result.getError());
        }
        return result.mapError(detail -> new OOError(errorTitle, messageOnFailureToObtain));
    }

    private static OOVoidResult<OOError> checkRangeOverlaps(XTextDocument doc, OOFrontend frontend) {
        final String errorTitle = "Overlapping ranges";
        boolean requireSeparation = false;
        int maxReportedOverlaps = 10;
        try {
            return frontend.checkRangeOverlaps(doc,
                                   new ArrayList<>(),
                                   requireSeparation,
                                   maxReportedOverlaps)
                           .mapError(OOError::from);
        } catch (NoDocumentException ex) {
            return OOVoidResult.error(OOError.from(ex).setTitle(errorTitle));
        } catch (WrappedTargetException ex) {
            return OOVoidResult.error(OOError.fromMisc(ex).setTitle(errorTitle));
        }
    }

    private static OOVoidResult<OOError> checkRangeOverlapsWithCursor(XTextDocument doc, OOFrontend frontend) {
        final String errorTitle = "Ranges overlapping with cursor";

        List<RangeForOverlapCheck<CitationGroupId>> userRanges;
        userRanges = frontend.viewCursorRanges(doc);

        boolean requireSeparation = false;
        OOVoidResult<JabRefException> res;
        try {
            res = frontend.checkRangeOverlapsWithCursor(doc,
                    userRanges,
                    requireSeparation);
        } catch (NoDocumentException ex) {
            return OOVoidResult.error(OOError.from(ex).setTitle(errorTitle));
        } catch (WrappedTargetException ex) {
            return OOVoidResult.error(OOError.fromMisc(ex).setTitle(errorTitle));
        }

        if (res.isError()) {
            final String xtitle = Localization.lang("The cursor is in a protected area.");
            return OOVoidResult.error(
                    new OOError(xtitle, xtitle + "\n" + res.getError().getLocalizedMessage() + "\n"));
        }
        return res.mapError(OOError::from);
    }

    /* ******************************************************
     *
     * Tests for preconditions.
     *
     * ******************************************************/

    private static OOVoidResult<OOError> checkIfOpenOfficeIsRecordingChanges(XTextDocument doc) {
        String errorTitle = Localization.lang("Recording and/or Recorded changes");
        try {
            boolean recordingChanges = UnoRedlines.getRecordChanges(doc);
            int nRedlines = UnoRedlines.countRedlines(doc);
            if (recordingChanges || (nRedlines > 0)) {
                String msg = "";
                if (recordingChanges) {
                    msg += Localization.lang("Cannot work with"
                            + " [Edit]/[Track Changes]/[Record] turned on.");
                }
                if (nRedlines > 0) {
                    if (recordingChanges) {
                        msg += "\n";
                    }
                    msg += Localization.lang("Changes by JabRef"
                            + " could result in unexpected interactions with"
                            + " recorded changes.");
                    msg += "\n";
                    msg += Localization.lang("Use [Edit]/[Track Changes]/[Manage] to resolve them first.");
                }
                return OOVoidResult.error(new OOError(errorTitle, msg));
            }
        } catch (WrappedTargetException ex) {
            String msg = Localization.lang("Error while checking if Writer"
                    + " is recording changes or has recorded changes.");
            return OOVoidResult.error(new OOError(errorTitle, msg, ex));
        }
        return OOVoidResult.ok();
    }

    OOVoidResult<OOError> styleIsRequired(OOStyle style) {
        if (style == null) {
            return OOVoidResult.error(OOError.noValidStyleSelected());
        } else {
            return OOVoidResult.ok();
        }
    }

    OOResult<OOFrontend, OOError> getFrontend(XTextDocument doc) {
        final String errorTitle = "Unable to get frontend";
        try {
            return OOResult.ok(new OOFrontend(doc));
        } catch (NoDocumentException ex) {
            return OOResult.error(OOError.from(ex).setTitle(errorTitle));
        } catch (WrappedTargetException
                 | RuntimeException ex) {
            return OOResult.error(OOError.fromMisc(ex).setTitle(errorTitle));
        }
    }

    OOVoidResult<OOError> databaseIsRequired(List<BibDatabase> databases,
                                             Supplier<OOError> fun) {
        if (databases == null || databases.isEmpty()) {
            return OOVoidResult.error(fun.get());
        } else {
            return OOVoidResult.ok();
        }
    }

    OOVoidResult<OOError> selectedBibEntryIsRequired(List<BibEntry> entries,
                                                     Supplier<OOError> fun) {
        if (entries == null || entries.isEmpty()) {
            return OOVoidResult.error(fun.get());
        } else {
            return OOVoidResult.ok();
        }
    }

    /*
     * Checks existence and also checks if it is not an internal name.
     */
    private OOVoidResult<OOError> checkStyleExistsInTheDocument(String familyName,
                                                                String styleName,
                                                                XTextDocument doc,
                                                                String labelInJstyleFile,
                                                                String pathToStyleFile)
            throws
            WrappedTargetException {

        Optional<String> internalName = UnoStyle.getInternalNameOfStyle(doc, familyName, styleName);

        if (internalName.isEmpty()) {
            String msg =
                    switch (familyName) {
                        case UnoStyle.PARAGRAPH_STYLES ->
                                Localization.lang("The %0 paragraph style '%1' is missing from the document",
                                        labelInJstyleFile,
                                        styleName);
                        case UnoStyle.CHARACTER_STYLES ->
                                Localization.lang("The %0 character style '%1' is missing from the document",
                                        labelInJstyleFile,
                                        styleName);
                        default ->
                                throw new IllegalArgumentException("Expected " + UnoStyle.CHARACTER_STYLES
                                        + " or " + UnoStyle.PARAGRAPH_STYLES
                                        + " for familyName");
                    }
                            + "\n"
                            + Localization.lang("Please create it in the document or change in the file:")
                            + "\n"
                            + pathToStyleFile;
            return OOVoidResult.error(new OOError("StyleIsNotKnown", msg));
        }

        if (!internalName.get().equals(styleName)) {
            String msg =
                    switch (familyName) {
                        case UnoStyle.PARAGRAPH_STYLES ->
                                Localization.lang("The %0 paragraph style '%1' is a display name for '%2'.",
                                        labelInJstyleFile,
                                        styleName,
                                        internalName.get());
                        case UnoStyle.CHARACTER_STYLES ->
                                Localization.lang("The %0 character style '%1' is a display name for '%2'.",
                                        labelInJstyleFile,
                                        styleName,
                                        internalName.get());
                        default ->
                                throw new IllegalArgumentException("Expected " + UnoStyle.CHARACTER_STYLES
                                        + " or " + UnoStyle.PARAGRAPH_STYLES
                                        + " for familyName");
                    }
                            + "\n"
                            + Localization.lang("Please use the latter in the style file below"
                            + " to avoid localization problems.")
                            + "\n"
                            + pathToStyleFile;
            return OOVoidResult.error(new OOError("StyleNameIsNotInternal", msg));
        }
        return OOVoidResult.ok();
    }

    public OOVoidResult<OOError> checkStylesExistInTheDocument(JStyle jStyle, XTextDocument doc) {
        String pathToStyleFile = jStyle.getPath();

        List<OOVoidResult<OOError>> results = new ArrayList<>();
        try {
            results.add(checkStyleExistsInTheDocument(UnoStyle.PARAGRAPH_STYLES,
                    jStyle.getReferenceHeaderParagraphFormat(),
                    doc,
                    "ReferenceHeaderParagraphFormat",
                    pathToStyleFile));
            results.add(checkStyleExistsInTheDocument(UnoStyle.PARAGRAPH_STYLES,
                    jStyle.getReferenceParagraphFormat(),
                    doc,
                    "ReferenceParagraphFormat",
                    pathToStyleFile));
            if (jStyle.isFormatCitations()) {
                results.add(checkStyleExistsInTheDocument(UnoStyle.CHARACTER_STYLES,
                        jStyle.getCitationCharacterFormat(),
                        doc,
                        "CitationCharacterFormat",
                        pathToStyleFile));
            }
        } catch (WrappedTargetException ex) {
            results.add(OOVoidResult.error(new OOError("Other error in checkStyleExistsInTheDocument",
                    ex.getMessage(),
                    ex)));
        }

        return collectResults("checkStyleExistsInTheDocument failed", results);
    }

    /* ******************************************************
     *
     * ManageCitationsDialogView
     *
     * ******************************************************/
    public OOResult<List<CitationEntry>, OOError> getCitationEntries() {
        OOResult<XTextDocument, OOError> odoc = getXTextDocument();
        if (odoc.isError()) {
            return OOResult.error(odoc.getError());
        }
        XTextDocument doc = odoc.get();

        OOVoidResult<OOError> changesCheck = checkIfOpenOfficeIsRecordingChanges(doc);
        if (changesCheck.isError()) {
            return OOResult.error(changesCheck.getError());
        }

        try {
            return OOResult.ok(ManageCitations.getCitationEntries(doc));
        } catch (NoDocumentException ex) {
            return OOResult.error(OOError.from(ex));
        } catch (DisposedException ex) {
            return OOResult.error(OOError.from(ex));
        } catch (WrappedTargetException ex) {
            LOGGER.warn("getCitationEntries", ex);
            return OOResult.error(OOError.fromMisc(ex));
        }
    }

    /// Apply editable parts of citationEntries to the document: store pageInfo.
    ///
    /// Does not change presentation.
    ///
    /// Note: we use no undo context here, because only DocumentConnection.setUserDefinedStringPropertyValue() is called, and Undo in LO will not undo that.
    ///
    /// GUI: "Manage citations" dialog "OK" button. Called from: ManageCitationsDialogViewModel.storeSettings
    ///
    ///
    /// Currently the only editable part is pageInfo.
    ///
    /// Since the only call to applyCitationEntries() only changes pageInfo w.r.t those returned by getCitationEntries(), we can do with the following restrictions:
    ///
    /// -  Missing pageInfo means no action.
    /// -  Missing CitationEntry means no action (no attempt to remove
    /// citation from the text).
    public OOVoidResult<OOError> applyCitationEntries(List<CitationEntry> citationEntries) {
        OOResult<XTextDocument, OOError> odoc = getXTextDocument();
        if (odoc.isError()) {
            return OOVoidResult.error(odoc.getError());
        }
        XTextDocument doc = odoc.get();

        try {
            ManageCitations.applyCitationEntries(doc, citationEntries);
            return OOVoidResult.ok();
        } catch (NoDocumentException ex) {
            return OOVoidResult.error(OOError.from(ex));
        } catch (DisposedException ex) {
            return OOVoidResult.error(OOError.from(ex));
        } catch (PropertyVetoException
                 | IllegalTypeException
                 | WrappedTargetException
                 | com.sun.star.lang.IllegalArgumentException ex) {
            LOGGER.warn("applyCitationEntries", ex);
            return OOVoidResult.error(OOError.fromMisc(ex));
        }
    }

    /// Creates a citation group from `entries` at the cursor.
    ///
    /// Uses LO undo context "Insert citation".
    ///
    /// Note: Undo does not remove or reestablish custom properties.
    ///
    /// Consistency: for each entry in `entries`: looking it up in `syncOptions.get().databases` (if present) should yield `database`.
    ///
    /// @param entries            The entries to cite.
    /// @param bibDatabaseContext The database the entries belong to (all of them). Used when creating the citation mark.
    /// @param style              The bibliography style we are using.
    /// @param citationType       Indicates whether it is an in-text citation, a citation in parenthesis or an invisible citation.
    /// @param pageInfo           A single page-info for these entries. Attributed to the last entry.
    /// @param syncOptions        Indicates whether in-text citations should be refreshed in the document. Optional.empty() indicates no refresh. Otherwise, provides options for refreshing the reference list.
    public OOVoidResult<OOError> insertEntry(List<BibEntry> entries,
                                             BibDatabaseContext bibDatabaseContext,
                                             BibEntryTypesManager bibEntryTypesManager,
                                             OOStyle style,
                                             CitationType citationType,
                                             String pageInfo,
                                             Optional<Update.SyncOptions> syncOptions) {

        OOResult<XTextDocument, OOError> odoc = getXTextDocument();
        if (odoc.isError()) {
            return odoc.asVoidResult();
        }
        XTextDocument doc = odoc.get();

        OOVoidResult<OOError> preconditions = collectResults("",
                List.of(
                        styleIsRequired(style),
                        selectedBibEntryIsRequired(entries, OOError::noEntriesSelectedForCitation)
                ));
        if (preconditions.isError()) {
            return preconditions;
        }

        OOResult<OOFrontend, OOError> frontend = getFrontend(doc);
        if (frontend.isError()) {
            return frontend.asVoidResult();
        }

        OOResult<XTextCursor, OOError> cursor = getUserCursorForTextInsertion(doc, "");
        if (cursor.isError()) {
            return cursor.asVoidResult();
        }

        OOVoidResult<OOError> overlapCheck = checkRangeOverlapsWithCursor(doc, frontend.get());
        if (overlapCheck.isError()) {
            return overlapCheck;
        }

        if (style instanceof JStyle jStyle) {
            OOVoidResult<OOError> styleChecks = collectResults("",
                    List.of(
                            checkStylesExistInTheDocument(jStyle, doc),
                            checkIfOpenOfficeIsRecordingChanges(doc)
                    ));
            if (styleChecks.isError()) {
                return styleChecks;
            }
        }

        /*
         * For sync we need a FunctionalTextViewCursor and an open database.
         */
        OOResult<FunctionalTextViewCursor, OOError> fcursor = null;
        if (syncOptions.isPresent()) {
            fcursor = getFunctionalTextViewCursor(doc, "");
            syncOptions.map(e -> e.setAlwaysAddCitedOnPages(openOfficePreferences.getAlwaysAddCitedOnPages()));
            if (fcursor.isError()) {
                return fcursor.asVoidResult();
            }
            OOVoidResult<OOError> dbCheck = databaseIsRequired(syncOptions.get().databases,
                    OOError::noDataBaseIsOpenForSyncingAfterCitation);
            if (dbCheck.isError()) {
                return dbCheck;
            }
        }

        try {
            UnoUndo.enterUndoContext(doc, "Insert citation");
            if (style instanceof CitationStyle citationStyle) {
                insertCSLCitation(entries,
                        doc,
                        citationType,
                        citationStyle,
                        bibDatabaseContext,
                        bibEntryTypesManager,
                        cursor,
                        syncOptions);
            } else if (style instanceof JStyle jStyle) {
                insertJStyleCitation(entries,
                        doc,
                        citationType,
                        jStyle,
                        frontend,
                        cursor,
                        bibDatabaseContext,
                        syncOptions,
                        pageInfo,
                        fcursor);
            }
            return OOVoidResult.ok();
        } catch (NoDocumentException ex) {
            return OOVoidResult.error(OOError.from(ex));
        } catch (DisposedException ex) {
            return OOVoidResult.error(OOError.from(ex));
        } catch (CreationException
                 | WrappedTargetException
                 | PropertyVetoException
                 | IllegalTypeException
                 | NotRemoveableException ex) {
            LOGGER.warn("insertEntry", ex);
            return OOVoidResult.error(OOError.fromMisc(ex));
        } catch (com.sun.star.uno.Exception ex) {
            LOGGER.error("insertEntry", ex);
            return OOVoidResult.error(OOError.fromMisc(ex));
        } finally {
            UnoUndo.leaveUndoContext(doc);
        }
    }

    /// Helper method for guiActionInsertEntry. Handles CSL citation insertion
    /// Throws CreationException, com.sun.star.uno.Exception
    /// Caught by guiActionInsertEntry
    ///
    /// @param entries            The entries to cite.
    /// @param bibDatabaseContext The database the entries belong to (all of them). Used when creating the citation mark.
    /// @param citationType       Indicates whether it is an in-text citation, a citation in parenthesis or an invisible citation.
    /// @param citationStyle      Indicates style, name and path of citation
    /// @param syncOptions        Indicates whether in-text citations should be refreshed in the document. Optional.empty() indicates no refresh. Otherwise, provides options for refreshing the reference list.
    public void insertCSLCitation(List<BibEntry> entries,
                                  XTextDocument doc,
                                  CitationType citationType,
                                  CitationStyle citationStyle,
                                  BibDatabaseContext bibDatabaseContext,
                                  BibEntryTypesManager bibEntryTypesManager,
                                  OOResult<XTextCursor,
                                          OOError> cursor,
                                  Optional<Update.SyncOptions> syncOptions)
            throws CreationException, com.sun.star.uno.Exception {
        try {
            // Lock document controllers - disable refresh during the process (avoids document flicker during writing)
            // MUST always be paired with an unlockControllers() call
            doc.lockControllers();

            if (citationType == CitationType.AUTHORYEAR_PAR) {
                // "Cite" button
                cslCitationOOAdapter.insertCitation(cursor.get(), citationStyle, entries, bibDatabaseContext, bibEntryTypesManager);
            } else if (citationType == CitationType.AUTHORYEAR_INTEXT) {
                // "Cite in-text" button
                cslCitationOOAdapter.insertInTextCitation(cursor.get(), citationStyle, entries, bibDatabaseContext, bibEntryTypesManager);
            } else if (citationType == CitationType.INVISIBLE_CIT) {
                // "Insert empty citation"
                cslCitationOOAdapter.insertEmptyCitation(cursor.get(), citationStyle, entries);
            }

            // If "Automatically sync bibliography when inserting citations" is enabled
            if (citationStyle.hasBibliography()) {
                syncOptions.ifPresent(options -> updateDocument(options.databases, citationStyle));
            }
        } finally {
            // Release controller lock
            doc.unlockControllers();
        }
    }

    /// Helper method for guiActionInsertEntry
    /// Throws PropertyVetoException, WrappedTargetException, IllegalTypeException, NotRemoveableException, CreationException, NoDocumentException
    /// Exceptions caught by guiActionInsertEntry
    ///
    /// @param entries            The entries to cite.
    /// @param citationType       Indicates whether it is an in-text citation, a citation in parentheses or an invisible citation.
    /// @param jStyle             Indicates citation formating in JStyle
    /// @param bibDatabaseContext The database the entries belong to (all of them). Used when creating the citation mark.
    /// @param syncOptions        Indicates whether in-text citations should be refreshed in the document. Optional.empty() indicates no refresh. Otherwise, provides options for refreshing the reference list.
    /// @param pageInfo           A single page-info for these entries. Attributed to the last entry.
    public void insertJStyleCitation(List<BibEntry> entries,
                                     XTextDocument doc,
                                     CitationType citationType,
                                     JStyle jStyle,
                                     OOResult<OOFrontend,
                                             OOError> frontend,
                                     OOResult<XTextCursor,
                                             OOError> cursor,
                                     BibDatabaseContext bibDatabaseContext,
                                     Optional<Update.SyncOptions> syncOptions,
                                     String pageInfo,
                                     OOResult<FunctionalTextViewCursor,
                                             OOError> fcursor)
            throws PropertyVetoException,
            WrappedTargetException,
            IllegalTypeException,
            NotRemoveableException,
            CreationException,
            NoDocumentException {
        EditInsert.insertCitationGroup(doc,
                frontend.get(),
                cursor.get(),
                entries,
                bibDatabaseContext.getDatabase(),
                jStyle,
                citationType,
                pageInfo);

        if (syncOptions.isPresent()) {
            Update.resyncDocument(doc, jStyle, fcursor.get(), syncOptions.get());
        }
    }

    /// Merges citation groups in the document.
    ///
    /// Uses LO undo context "Merge citations".
    public OOVoidResult<OOError> mergeCitationGroups(List<BibDatabase> databases, OOStyle style) {
        if (!(style instanceof JStyle jStyle)) {
            return OOVoidResult.ok();
        }

        OOResult<XTextDocument, OOError> odoc = getXTextDocument();
        if (odoc.isError()) {
            return odoc.asVoidResult();
        }

        OOVoidResult<OOError> preconditions = collectResults("",
                List.of(
                        styleIsRequired(jStyle),
                        databaseIsRequired(databases, OOError::noDataBaseIsOpen)
                ));
        if (preconditions.isError()) {
            return preconditions;
        }
        XTextDocument doc = odoc.get();

        OOResult<FunctionalTextViewCursor, OOError> fcursor = getFunctionalTextViewCursor(doc, "");
        OOVoidResult<OOError> checks = collectResults("",
                List.of(
                        fcursor.asVoidResult(),
                        checkStylesExistInTheDocument(jStyle, doc),
                        checkIfOpenOfficeIsRecordingChanges(doc)
                ));
        if (checks.isError()) {
            return checks;
        }

        try {
            UnoUndo.enterUndoContext(doc, "Merge citations");

            OOFrontend frontend = new OOFrontend(doc);
            boolean madeModifications = EditMerge.mergeCitationGroups(doc, frontend, jStyle);
            if (madeModifications) {
                UnoCrossRef.refresh(doc);
                Update.SyncOptions syncOptions = new Update.SyncOptions(databases);
                Update.resyncDocument(doc, jStyle, fcursor.get(), syncOptions);
            }
            return OOVoidResult.ok();
        } catch (NoDocumentException ex) {
            return OOVoidResult.error(OOError.from(ex));
        } catch (DisposedException ex) {
            return OOVoidResult.error(OOError.from(ex));
        } catch (CreationException
                 | IllegalTypeException
                 | NotRemoveableException
                 | PropertyVetoException
                 | WrappedTargetException
                 | com.sun.star.lang.IllegalArgumentException ex) {
            LOGGER.warn("mergeCitationGroups", ex);
            return OOVoidResult.error(OOError.fromMisc(ex));
        } finally {
            UnoUndo.leaveUndoContext(doc);
            if (fcursor != null && !fcursor.isError()) {
                fcursor.get().restore(doc);
            }
        }
    }

    /// GUI action "Separate citations".
    ///
    /// Do the opposite of MergeCitationGroups. Combined markers are split, with a space inserted between.
    /// Do the opposite of mergeCitationGroups. Combined markers are split, with a space inserted between.
    ///
    /// Uses LO undo context "Separate citations".
    public OOVoidResult<OOError> separateCitations(List<BibDatabase> databases, OOStyle style) {
        if (!(style instanceof JStyle jStyle)) {
            return OOVoidResult.ok();
        }

        OOResult<XTextDocument, OOError> odoc = getXTextDocument();
        if (odoc.isError()) {
            return odoc.asVoidResult();
        }

        OOVoidResult<OOError> preconditions = collectResults("",
                List.of(
                        styleIsRequired(jStyle),
                        databaseIsRequired(databases, OOError::noDataBaseIsOpen)
                ));
        if (preconditions.isError()) {
            return preconditions;
        }
        XTextDocument doc = odoc.get();

        OOResult<FunctionalTextViewCursor, OOError> fcursor = getFunctionalTextViewCursor(doc, "");
        OOVoidResult<OOError> checks = collectResults("",
                List.of(
                        fcursor.asVoidResult(),
                        checkStylesExistInTheDocument(jStyle, doc),
                        checkIfOpenOfficeIsRecordingChanges(doc)
                ));
        if (checks.isError()) {
            return checks;
        }

        try {
            UnoUndo.enterUndoContext(doc, "Separate citations");

            OOFrontend frontend = new OOFrontend(doc);
            boolean madeModifications = EditSeparate.separateCitations(doc, frontend, databases, jStyle);
            if (madeModifications) {
                UnoCrossRef.refresh(doc);
                Update.SyncOptions syncOptions = new Update.SyncOptions(databases);
                Update.resyncDocument(doc, jStyle, fcursor.get(), syncOptions);
            }
            return OOVoidResult.ok();
        } catch (NoDocumentException ex) {
            return OOVoidResult.error(OOError.from(ex));
        } catch (DisposedException ex) {
            return OOVoidResult.error(OOError.from(ex));
        } catch (CreationException
                 | IllegalTypeException
                 | NotRemoveableException
                 | PropertyVetoException
                 | WrappedTargetException
                 | com.sun.star.lang.IllegalArgumentException ex) {
            LOGGER.warn("separateCitations", ex);
            return OOVoidResult.error(OOError.fromMisc(ex));
        } finally {
            UnoUndo.leaveUndoContext(doc);
            if (fcursor != null && !fcursor.isError()) {
                fcursor.get().restore(doc);
            }
        }
    }

    /// GUI action for "Export cited"
    ///
    /// Does not refresh the bibliography.
    ///
    /// @param returnPartialResult If there are some unresolved keys, shall we return an otherwise nonempty result, or Optional.empty()?
    /// Lógica pura: exporta as referências citadas para um novo BibDatabase.
    ///
    /// Não atualiza a bibliografia.
    ///
    /// @param returnPartialResult Se houver chaves não resolvidas, retorna o
    ///                            resultado parcial (true) ou erro (false).
    public OOResult<BibDatabase, OOError> exportCitedHelper(
            List<BibDatabase> databases,
            boolean returnPartialResult) {

        OOResult<XTextDocument, OOError> odoc = getXTextDocument();
        if (odoc.isError()) {
            return OOResult.error(odoc.getError());
        }

        OOVoidResult<OOError> dbCheck = databaseIsRequired(databases, OOError::noDataBaseIsOpenForExport);
        if (dbCheck.isError()) {
            return OOResult.error(dbCheck.getError());
        }

        XTextDocument doc = odoc.get();

        try {
            ExportCited.GenerateDatabaseResult result;
            try {
                UnoUndo.enterUndoContext(doc, "Changes during \"Export cited\"");
                result = ExportCited.generateDatabase(doc, databases);
            } finally {
                // There should be no changes, thus no Undo entry should appear
                // in LibreOffice.
                UnoUndo.leaveUndoContext(doc);
            }

            if (!result.newDatabase.hasEntries()) {
                return OOResult.error(new OOError(
                        "",
                        Localization.lang("Your OpenOffice/LibreOffice document references"
                                + " no citation keys"
                                + " which could also be found in your current library.")));
            }

            List<String> unresolvedKeys = result.unresolvedKeys;
            if (!unresolvedKeys.isEmpty()) {
                OOError unresolvedError = new OOError(
                        "",
                        Localization.lang(
                                "Your OpenOffice/LibreOffice document references"
                                        + " at least %0 citation keys"
                                        + " which could not be found in your current library."
                                        + " Some of these are %1.",
                                String.valueOf(unresolvedKeys.size()),
                                String.join(", ", unresolvedKeys)));
                if (returnPartialResult) {
                    return OOResult.ok(result.newDatabase);
                } else {
                    return OOResult.error(unresolvedError);
                }
            }

            return OOResult.ok(result.newDatabase);

        } catch (NoDocumentException ex) {
            return OOResult.error(OOError.from(ex));
        } catch (DisposedException ex) {
            return OOResult.error(OOError.from(ex));
        } catch (WrappedTargetException
                 | com.sun.star.lang.IllegalArgumentException ex) {
            LOGGER.warn("Problem generating new database.", ex);
            return OOResult.error(OOError.fromMisc(ex));
        }
    }

    /// Refreshes citation markers and bibliography.
    ///
    /// @param databases Must have at least one.
    /// @param style     Style.
    public OOVoidResult<OOError> updateDocument(List<BibDatabase> databases, OOStyle style) {

        OOResult<XTextDocument, OOError> odoc = getXTextDocument();
        if (odoc.isError()) {
            return odoc.asVoidResult();
        }

        OOVoidResult<OOError> preconditions = styleIsRequired(style);
        if (preconditions.isError()) {
            return preconditions;
        }
        XTextDocument doc = odoc.get();

        OOResult<FunctionalTextViewCursor, OOError> fcursor = getFunctionalTextViewCursor(doc, "");

        if (style instanceof JStyle jStyle) {
            OOVoidResult<OOError> checks = collectResults("",
                    List.of(
                            fcursor.asVoidResult(),
                            checkStylesExistInTheDocument(jStyle, doc),
                            checkIfOpenOfficeIsRecordingChanges(doc)
                    ));
            if (checks.isError()) {
                return checks;
            }

            try {
                OOFrontend frontend = new OOFrontend(doc);

                OOVoidResult<OOError> overlapCheck = checkRangeOverlaps(doc, frontend);
                if (overlapCheck.isError()) {
                    return overlapCheck;
                }

                updateJStyleBibliography(databases, jStyle, doc, frontend, fcursor, "");
                return OOVoidResult.ok();
            } catch (NoDocumentException ex) {
                return OOVoidResult.error(OOError.from(ex));
            } catch (DisposedException ex) {
                return OOVoidResult.error(OOError.from(ex));
            } catch (CreationException
                     | WrappedTargetException
                     | com.sun.star.lang.IllegalArgumentException ex) {
                LOGGER.warn("updateDocument: JStyle", ex);
                return OOVoidResult.error(OOError.fromMisc(ex));
            }

        } else if (style instanceof CitationStyle citationStyle) {
            if (!citationStyle.hasBibliography()) {
                return OOVoidResult.ok();
            }

            OOVoidResult<OOError> checks = collectResults("",
                    List.of(
                            fcursor.asVoidResult(),
                            checkIfOpenOfficeIsRecordingChanges(doc)
                    ));
            if (checks.isError()) {
                return checks;
            }

            try {
                updateCSLBibliography(databases, citationStyle, doc, fcursor);
                return OOVoidResult.ok();
            } catch (CreationException
                     | com.sun.star.lang.IllegalArgumentException ex) {
                LOGGER.error("updateDocument: CSL", ex);
                return OOVoidResult.error(OOError.fromMisc(ex));
            }
        }

        return OOVoidResult.ok();
    }

    /// Helper method for guiActionUpdateDocument, refreshes a JStyle bibliography.
    ///
    /// @param databases        Must have at least one.
    /// @param jStyle           Indicates citation formating in JStyle.
    /// @param doc              Text document.
    /// @param frontend,fcursor Used to synchronize document.
    /// @param errorTitle       Error message for user.
    private void updateJStyleBibliography(List<BibDatabase> databases, JStyle jStyle, XTextDocument doc, OOFrontend frontend,
                                          OOResult<FunctionalTextViewCursor, OOError> fcursor, String errorTitle)
            throws CreationException, NoDocumentException, WrappedTargetException {
        List<String> unresolvedKeys;
        try {
            UnoUndo.enterUndoContext(doc, "Refresh bibliography");

            Update.SyncOptions syncOptions = new Update.SyncOptions(databases);
            syncOptions
                    .setUpdateBibliography(true)
                    .setAlwaysAddCitedOnPages(openOfficePreferences.getAlwaysAddCitedOnPages());

            unresolvedKeys = Update.synchronizeDocument(doc, frontend, jStyle, fcursor.get(), syncOptions);
        } finally {
            UnoUndo.leaveUndoContext(doc);
            fcursor.get().restore(doc);
        }

        if (!unresolvedKeys.isEmpty()) {
            String msg = Localization.lang(
                    "Your OpenOffice/LibreOffice document references the citation key '%0',"
                            + " which could not be found in your current library.",
                    unresolvedKeys.getFirst());
            dialogService.showErrorDialogAndWait(errorTitle, msg);
        }
    }

    /// Helper method for guiActionUpdateDocument, refreshes a CSL bibliography.
    ///
    /// @param databases     Must have at least one.
    /// @param citationStyle Citation style to update bibliography with.
    /// @param doc           Text document.
    /// @param fcursor       Used to synchronize document.
    private void updateCSLBibliography(List<BibDatabase> databases, CitationStyle citationStyle, XTextDocument doc,
                                       OOResult<FunctionalTextViewCursor, OOError> fcursor)
            throws CreationException {
        try {
            UnoUndo.enterUndoContext(doc, "Create CSL bibliography");

            // Collect only cited entries from all databases
            List<BibEntry> citedEntries = databases.stream()
                                                   .flatMap(database -> database.getEntries().stream())
                                                   .filter(cslCitationOOAdapter::isCitedEntry)
                                                   .collect(Collectors.toCollection(ArrayList::new));

            // If no entries are cited, show a message and return
            if (citedEntries.isEmpty()) {
                dialogService.showInformationDialogAndWait(
                        Localization.lang("Bibliography"),
                        Localization.lang("No cited entries found in the document.")
                );
                return;
            }

            // A separate database and database context
            BibDatabase bibDatabase = new BibDatabase(citedEntries);
            BibDatabaseContext bibDatabaseContext = new BibDatabaseContext(bibDatabase);

            // Lock document controllers - disable refresh during the process (avoids document flicker during writing)
            // MUST always be paired with an unlockControllers() call
            doc.lockControllers();

            cslUpdateBibliography.rebuildCSLBibliography(doc, cslCitationOOAdapter, citedEntries, citationStyle, bibDatabaseContext, Injector.instantiateModelOrService(BibEntryTypesManager.class));
        } catch (NoDocumentException | com.sun.star.uno.Exception e) {
            dialogService.notify(Localization.lang("No document found or LibreOffice insertion failure"));
            LOGGER.error("Could not update CSL bibliography", e);
        } finally {
            doc.unlockControllers();
            UnoUndo.leaveUndoContext(doc);
            fcursor.get().restore(doc);
        }
    }
}
