package org.jabref.gui.openoffice;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;

import org.jabref.gui.DialogService;
import org.jabref.logic.l10n.Localization;
import org.jabref.logic.openoffice.NoDocumentFoundException;
import org.jabref.logic.openoffice.OpenOfficePreferences;
import org.jabref.logic.openoffice.action.Update;
import org.jabref.logic.openoffice.style.OOStyle;
import org.jabref.model.database.BibDatabase;
import org.jabref.model.database.BibDatabaseContext;
import org.jabref.model.entry.BibEntry;
import org.jabref.model.entry.BibEntryTypesManager;
import org.jabref.model.openoffice.CitationEntry;
import org.jabref.model.openoffice.style.CitationType;
import org.jabref.model.openoffice.uno.CreationException;
import org.jabref.model.openoffice.util.OOResult;
import org.jabref.model.openoffice.util.OOVoidResult;

import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.WrappedTargetException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class OOBibBaseGUI {

    private static final Logger LOGGER = LoggerFactory.getLogger(OOBibBaseGUI.class);

    private final OOBibBase logic;

    private final DialogService dialogService;

    public OOBibBaseGUI(Path loPath, DialogService dialogService, OpenOfficePreferences openOfficePreferences)
            throws BootstrapException, CreationException, IOException, InterruptedException {
        this.logic = new OOBibBase(loPath, dialogService, openOfficePreferences);
        this.dialogService = dialogService;
    }

    OOBibBaseGUI(OOBibBase logic, DialogService dialogService) {
        this.logic = logic;
        this.dialogService = dialogService;
    }

    void showDialog(OOError err) {
        err.showErrorDialog(dialogService);
    }

    boolean testDialog(OOVoidResult<OOError> res) {
        return res.ifError(ex -> ex.showErrorDialog(dialogService)).isError();
    }

    boolean testDialog(String errorTitle, OOVoidResult<OOError> res) {
        return res.ifError(e -> showDialog(e.setTitle(errorTitle))).isError();
    }

    @SafeVarargs
    final boolean testDialog(String errorTitle, OOVoidResult<OOError>... results) {
        List<OOVoidResult<OOError>> resultList = Arrays.asList(results);
        return testDialog(logic.collectResults(errorTitle, resultList));
    }

    /// GUI action to select a Writer document to work with.
    ///
    /// @param autoSelectForSingle If true and only one document is available, select it automatically.
    public void selectDocument(boolean autoSelectForSingle) throws WrappedTargetException, NoSuchElementException {
        final String errorTitle = Localization.lang("Problem connecting");

        try {
            logic.getConnection().selectDocument(autoSelectForSingle);
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
            logic.initializeCitationAdapter(logic.getXTextDocument().get());
            dialogService.notify(Localization.lang("Connected to document") + ": "
                    + getCurrentDocumentTitle().orElse(""));
        }
    }

    /// A simple test for document availability.
    public boolean isConnectedToDocument() {
        return logic.isConnectedToDocument();
    }

    /// Check if the document connection is missing.
    public boolean isDocumentConnectionMissing() {
        return logic.isDocumentConnectionMissing();
    }

    /// The title of the current document, or Optional.empty()
    public Optional<String> getCurrentDocumentTitle() {
        return logic.getCurrentDocumentTitle();
    }

    /// Close any open office connection, if none exists does nothing
    public static void closeOfficeConnection() {
        OOBibBaseConnect.closeOfficeConnection();
    }

    public Optional<List<CitationEntry>> getCitationEntries() {
        final String errorTitle = Localization.lang("Problem collecting citations");

        OOResult<List<CitationEntry>, OOError> result = logic.getCitationEntries();
        if (testDialog(errorTitle, result.asVoidResult())) {
            return Optional.empty();
        }
        return Optional.of(result.get());
    }

    public void applyCitationEntries(List<CitationEntry> citationEntries) {
        final String errorTitle = Localization.lang("Problem modifying citation");
        testDialog(errorTitle, logic.applyCitationEntries(citationEntries));
    }

    public void insertEntry(List<BibEntry> entries,
                                     BibDatabaseContext bibDatabaseContext,
                                     BibEntryTypesManager bibEntryTypesManager,
                                     OOStyle style,
                                     CitationType citationType,
                                     String pageInfo,
                                     Optional<Update.SyncOptions> syncOptions) {

        final String errorTitle = "Could not insert citation";
        testDialog(errorTitle, logic.insertEntry(entries,
                bibDatabaseContext,
                bibEntryTypesManager,
                style,
                citationType,
                pageInfo,
                syncOptions));
    }

    public void mergeCitationGroups(List<BibDatabase> databases, OOStyle style) {
        final String errorTitle = Localization.lang("Problem combining cite markers");
        testDialog(errorTitle, logic.mergeCitationGroups(databases, style));
    }


    /// Do the opposite of MergeCitationGroups. Combined markers are split, with a space inserted between.
    public void separateCitations(List<BibDatabase> databases, OOStyle style) {
        final String errorTitle = Localization.lang("Problem during separating cite markers");
        testDialog(errorTitle, logic.separateCitations(databases, style));
    }

    /// Refreshes citation markers and bibliography.
    ///
    /// @param databases Must have at least one.
    public void updateDocument(List<BibDatabase> databases, OOStyle style) {
        final String errorTitle = Localization.lang("Unable to synchronize bibliography");

        OOBibBase.UpdateDocumentResult result = logic.updateDocument(databases, style);

        // Handle errors first
        if (result.isError()) {
            showDialog(result.getError().setTitle(errorTitle));
            return;
        }

        // Handle CSL-specific results
        if (result.cslResult().isPresent()) {
            switch (result.cslResult().get()) {
                case NO_CITED_ENTRIES -> dialogService.showInformationDialogAndWait(
                        Localization.lang("Bibliography"),
                        Localization.lang("No cited entries found in the document."));
                case ERROR -> dialogService.notify(
                        Localization.lang("No document found or LibreOffice insertion failure"));
                case SUCCESS -> {
                    // Success - nothing to show
                }
            }
        }
    }

    /// GUI action for "Export cited"
    ///
    /// Does not refresh the bibliography.
    ///
    /// @param returnPartialResult If there are some unresolved keys, shall we return an otherwise nonempty result, or Optional.empty()?
    public Optional<BibDatabase> exportCitedHelper(
            List<BibDatabase> databases,
            boolean returnPartialResult) {

        final String errorTitle = Localization.lang("Unable to generate new library");

        OOResult<BibDatabase, OOError> result =
                logic.exportCitedHelper(databases, returnPartialResult);

        if (testDialog(errorTitle, result.asVoidResult())) {
            return Optional.empty();
        }
        return Optional.of(result.get());
    }
}
