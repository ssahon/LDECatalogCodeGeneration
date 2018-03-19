//------------------------------------------------------------------------------
// <copyright file="Command1.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Document = Microsoft.CodeAnalysis.Document;

namespace VSIXEthos
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command1
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("7bb1c54f-85de-4d26-988e-b01627a59bf7");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command1"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private Command1(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new OleMenuCommand(this.MenuItemCallback, menuCommandID);
                menuItem.BeforeQueryStatus += MenuItem_BeforeQueryStatus;
                commandService.AddCommand(menuItem);
            }
        }

        private void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand myCommand = (OleMenuCommand)sender;
            IVsMonitorSelection monitorSelection = ServiceProvider.GetService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            uint solutionExistCookie;
            int pfActive;
            Guid NosolutionGuid = VSConstants.UICONTEXT_SolutionExists;
            monitorSelection.GetCmdUIContextCookie(ref NosolutionGuid, out solutionExistCookie);
            monitorSelection.IsCmdUIContextActive(solutionExistCookie, out pfActive);
            if (pfActive == 1 && GetActiveTextView() != null)
            {
                myCommand.Visible = true;
            }
            else
            {
                myCommand.Visible = false;
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command1 Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
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
        public static void Initialize(Package package)
        {
            Instance = new Command1(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var v = GetActiveTextView();
            var textView = GetTextViewFromVsTextView(v);

            Microsoft.VisualStudio.Text.SnapshotPoint caretPosition = textView.Caret.Position.BufferPosition;
            
            Document document = caretPosition.Snapshot.GetOpenDocumentInCurrentContextWithChanges();

            var method =
                document.GetSyntaxRootAsync().Result.FindToken(caretPosition).Parent.AncestorsAndSelf().
                    OfType<Microsoft.CodeAnalysis.CSharp.Syntax.MethodDeclarationSyntax>().FirstOrDefault();
            if (method != null)
            {
                SemanticModel semanticModel = document.GetSemanticModelAsync().Result;

                var methodSymbol = semanticModel.GetDeclaredSymbol(method);

                if (methodSymbol.IsStatic)
                {
                    VsShellUtilities.ShowMessageBox(this.ServiceProvider, "Cannot use static method", "Error",
                    OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    return;
                }

                if (methodSymbol.DeclaredAccessibility != Accessibility.Public)
                {
                    VsShellUtilities.ShowMessageBox(this.ServiceProvider, "Cannot use non public method", "Error",
                    OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    return;
                }

                var symbolDisplayFormat = new SymbolDisplayFormat(
                    typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces,
                    genericsOptions: SymbolDisplayGenericsOptions.IncludeTypeParameters,
                    parameterOptions: SymbolDisplayParameterOptions.IncludeType);

                string returnTypeName = methodSymbol.ReturnType.ToDisplayString(symbolDisplayFormat);

                StringBuilder sb = new StringBuilder();

                Trace.WriteLine($"method name: {methodSymbol.Name}");
                sb.AppendLine($"method name: {methodSymbol.Name}");
                XElement root = new XElement("Step");
                root.Add(new XAttribute("type", "StepCode"));
                root.Add(new XAttribute("method", methodSymbol.Name));

                if (method.ParameterList.Parameters.Count > 0)
                {
                    XElement inputs = new XElement("Inputs");

                    foreach (var parameter in method.ParameterList.Parameters)
                    {
                        var p = semanticModel.GetDeclaredSymbol(parameter);
                        string parameterTypeName = p.ToDisplayString(symbolDisplayFormat);
                        Trace.WriteLine($"input type: {parameterTypeName}");
                        sb.AppendLine($"input type: {parameterTypeName}");

                        XElement input = new XElement("Input");
                        input.Add(new XAttribute("name", p.Name));
                        input.Add(new XAttribute("type", parameterTypeName));
                        inputs.Add(input);
                    }

                    root.Add(inputs);
                }

                Trace.WriteLine($"return type: {returnTypeName}");
                sb.AppendLine($"return type: {returnTypeName}");

                if (returnTypeName != "System.Void")
                {
                    XElement output = new XElement("Output");
                    output.Add(new XAttribute("name", "result"));
                    output.Add(new XAttribute("type", returnTypeName));
                    root.Add(output);
                }

                CodeWindowViewModel vm = new CodeWindowViewModel();
                //vm.CatalogCode = sb.ToString();
                vm.CatalogCode = root.ToString();
                var w = new VSIXEthos.CodeWindow(vm);
                w.ShowModal();
            }
            else
            {
                VsShellUtilities.ShowMessageBox(this.ServiceProvider, "No method selected", "Error",
                    OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        /// <summary>  
        /// Find the active text view (if any) in the active document.  
        /// </summary>  
        /// <returns>The IVsTextView of the active view, or null if there is no active  
        /// document or the  
        /// active view in the active document is not a text view.</returns>  
        private IVsTextView GetActiveTextView()
        {
            IVsMonitorSelection selection =
                this.ServiceProvider.GetService(typeof(IVsMonitorSelection))
                                                    as IVsMonitorSelection;
            object frameObj = null;
            ErrorHandler.ThrowOnFailure(
                selection.GetCurrentElementValue(
                    (uint)VSConstants.VSSELELEMID.SEID_DocumentFrame, out frameObj));

            IVsWindowFrame frame = frameObj as IVsWindowFrame;
            if (frame == null)
            {
                return null;
            }

            return GetActiveView(frame);
        }

        private static IVsTextView GetActiveView(IVsWindowFrame windowFrame)
        {
            if (windowFrame == null)
            {
                throw new ArgumentException("windowFrame");
            }

            object pvar;
            ErrorHandler.ThrowOnFailure(
                windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out pvar));

            IVsTextView textView = pvar as IVsTextView;
            if (textView == null)
            {
                IVsCodeWindow codeWin = pvar as IVsCodeWindow;
                if (codeWin != null)
                {
                    ErrorHandler.ThrowOnFailure(codeWin.GetLastActiveView(out textView));
                }
            }
            return textView;
        }

        private static IWpfTextView GetTextViewFromVsTextView(IVsTextView view)
        {

            if (view == null)
            {
                throw new ArgumentNullException("view");
            }

            IVsUserData userData = view as IVsUserData;
            if (userData == null)
            {
                throw new InvalidOperationException();
            }

            object objTextViewHost;
            if (VSConstants.S_OK
                   != userData.GetData(Microsoft.VisualStudio.Editor.DefGuidList.guidIWpfTextViewHost,
                                       out objTextViewHost))
            {
                throw new InvalidOperationException();
            }

            IWpfTextViewHost textViewHost = objTextViewHost as IWpfTextViewHost;
            if (textViewHost == null)
            {
                throw new InvalidOperationException();
            }

            return textViewHost.TextView;
        }
    }

    public static class Extension
    {
        public static string GetFullMetadataName(this INamespaceOrTypeSymbol symbol)
        {
            ISymbol s = symbol;
            var sb = new StringBuilder(s.MetadataName);

            var last = s;
            s = s.ContainingSymbol;
            while (!IsRootNamespace(s))
            {
                if (s is ITypeSymbol && last is ITypeSymbol)
                {
                    sb.Insert(0, '+');
                }
                else
                {
                    sb.Insert(0, '.');
                }
                sb.Insert(0, s.MetadataName);
                s = s.ContainingSymbol;
            }

            return sb.ToString();
        }

        private static bool IsRootNamespace(ISymbol s)
        {
            return s is INamespaceSymbol && ((INamespaceSymbol)s).IsGlobalNamespace;
        }
    }
}
