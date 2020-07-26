using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.Text.Editor;
using System.Windows;
using System.Windows.Forms;
using Microsoft.VisualStudio.Text;
using System.Diagnostics;

namespace VSAirliner
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class AirlinerHungryBackspace
    {

        private readonly Regex trailingWhitespaceRegex = new Regex(@"^.*?(\s+)$");
        private readonly Regex whitespaceRegex = new Regex(@"^\s+$");

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0102;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("8bf2d3cd-5f4b-4158-989a-05d405c58789");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="AirlinerHungryBackspace"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private AirlinerHungryBackspace(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static AirlinerHungryBackspace Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in AirlinerCutToEol's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new AirlinerHungryBackspace(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var docInfoRes = this.GetDocInfo();
            if (!docInfoRes.HasValue)
            {
                return;
            }

            var docInfo = docInfoRes.Value;

            //
            // If there is currently a selection, just delete that and be done.
            //
            if (docInfo.AnchorPos != docInfo.ActivePos)
            {
                int start = Math.Min(docInfo.AnchorPos, docInfo.ActivePos);
                int end   = Math.Max(docInfo.AnchorPos, docInfo.ActivePos);

                using (var edit = docInfo.Snapshot.TextBuffer.CreateEdit())
                {
                    edit.Delete(start, end - start);
                    edit.Apply();
                }

                return;
            }

            //
            // From this point on, we will delete as many characters as
            // charsToDelete is set to.
            //
            
            // Move the cursor out of any virtual space.  Refresh docInfo in
            // case the cursor moved.
            docInfo.View.Caret.MoveTo(new VirtualSnapshotPoint(docInfo.Snapshot, docInfo.View.Selection.ActivePoint.Position));
            docInfo = this.GetDocInfo().Value;

            int charsToDelete = 0;
            int solPos = docInfo.CurLine.Start.Position;   // sol: start of line

            if (docInfo.ActivePos == solPos)
            {
                // We are at the beginning of the line.  Delete over all line
                // ending characters.
                bool done = false;
                int curPos = docInfo.ActivePos;

                do
                {
                    curPos--;
                    var curChar = docInfo.Snapshot.GetText(curPos, 1);
                    System.Diagnostics.Trace.WriteLine($"curChar: {curChar}");
                    if (curChar == "\n" || curChar == "\r" || this.whitespaceRegex.IsMatch(curChar))
                    {
                        charsToDelete++;
                        System.Diagnostics.Trace.WriteLine($"charsToDelete++");
                    }
                    else
                    {
                        done = true;
                        System.Diagnostics.Trace.WriteLine($"done");
                    }
                } while (!done);
            }
            else
            {
                // We are not at the beginning of the line.

                Span toSolSpan = new Span(solPos, docInfo.ActivePos - solPos);
                string toSolText = docInfo.Snapshot.GetText(toSolSpan);

                // If whitespace is to the left, we want to backspace over all of it.
                var match = this.trailingWhitespaceRegex.Match(toSolText);
                if (match.Success)
                {
                    charsToDelete = match.Groups[1].Length;
                }
                else
                {
                    // Text to the left of the cursor is not whitelspace.
                    // Perform a normal backspace.
                    charsToDelete = 1;
                }
            }
            
            using (var edit = docInfo.Snapshot.TextBuffer.CreateEdit())
            {
                edit.Delete(docInfo.ActivePos - charsToDelete, charsToDelete);
                edit.Apply();
            }
        }

        /// <summary>
        /// Convenience helper method for getting useful information about the
        /// document being edited.
        /// </summary>
        /// <returns>A data structure containing useful information.</returns>
        private Nullable<DocInfo> GetDocInfo()
        {
            IWpfTextView view = ProjectHelpers.GetCurentTextView();
            if (view == null)
                return null;

            ITextSnapshot snapshot = view.TextSnapshot;
            ITextSelection selection = view.Selection;
            ITextSnapshotLine curLine = snapshot.GetLineFromPosition(selection.ActivePoint.Position.Position);
            return new DocInfo(view, snapshot, selection, curLine);
        }

        //private void MessageBox(string message)
        //{
        //    VsShellUtilities.ShowMessageBox(
        //        this.package,
        //        message,
        //        "Debug Message",
        //        OLEMSGICON.OLEMSGICON_INFO,
        //        OLEMSGBUTTON.OLEMSGBUTTON_OK,
        //        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        //}
    }
}
