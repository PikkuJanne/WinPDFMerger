// Test-only native PowerShell host. Records what Write-Host actually renders.
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Security;

namespace WinPDFMerge.Tests {
    public sealed class RecordingConsoleRawUI : PSHostRawUserInterface {
        public override ConsoleColor BackgroundColor { get; set; }
        public override ConsoleColor ForegroundColor { get; set; }
        public override Size BufferSize { get; set; }
        public override Coordinates CursorPosition { get; set; }
        public override int CursorSize { get; set; }
        public override bool KeyAvailable { get { return false; } }
        public override Size MaxPhysicalWindowSize { get { return new Size(80,25); } }
        public override Size MaxWindowSize { get { return new Size(80,25); } }
        public override Coordinates WindowPosition { get; set; }
        public override Size WindowSize { get; set; }
        public override string WindowTitle { get; set; }
        public override void FlushInputBuffer() { }
        public override BufferCell[,] GetBufferContents(Rectangle rectangle) { throw new NotSupportedException(); }
        public override KeyInfo ReadKey(ReadKeyOptions options) { throw new NotSupportedException(); }
        public override void ScrollBufferContents(Rectangle source, Coordinates destination, Rectangle clip, BufferCell fill) { throw new NotSupportedException(); }
        public override void SetBufferContents(Coordinates origin, BufferCell[,] contents) { throw new NotSupportedException(); }
        public override void SetBufferContents(Rectangle rectangle, BufferCell fill) { throw new NotSupportedException(); }
    }
    public sealed class ConsoleWrite {
        public string Text;
        public ConsoleColor? Foreground;
        public bool NewLine;
    }
    public sealed class RecordingConsoleUI : PSHostUserInterface {
        public readonly List<ConsoleWrite> Writes = new List<ConsoleWrite>();
        public bool RefuseColors;
        private readonly RecordingConsoleRawUI rawUI = new RecordingConsoleRawUI();
        public override PSHostRawUserInterface RawUI { get { return rawUI; } }
        public override string ReadLine() { throw new NotSupportedException(); }
        public override SecureString ReadLineAsSecureString() { throw new NotSupportedException(); }
        public override void Write(string value) { }
        public override void Write(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value) {
            if (RefuseColors) throw new NotSupportedException("Synthetic host refuses color");
        }
        public override void WriteLine(string value) { }
        public override void WriteLine(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value) {
            if (RefuseColors) throw new NotSupportedException("Synthetic host refuses color");
        }
        public override void WriteErrorLine(string value) { WriteLine(value); }
        public override void WriteDebugLine(string value) { WriteLine(value); }
        public override void WriteVerboseLine(string value) { WriteLine(value); }
        public override void WriteWarningLine(string value) { WriteLine(value); }
        public override void WriteInformation(InformationRecord record) {
            var message = record.MessageData as HostInformationMessage;
            if (message == null) { WriteLine(record.ToString()); return; }
            if (RefuseColors && message.ForegroundColor.HasValue && message.ForegroundColor.Value != rawUI.ForegroundColor) throw new NotSupportedException("Synthetic host refuses color");
            Writes.Add(new ConsoleWrite { Text=message.Message, Foreground=message.ForegroundColor, NewLine=message.NoNewLine != true });
        }
        public override void WriteProgress(long sourceId, ProgressRecord record) { throw new NotSupportedException(); }
        public override Dictionary<string, PSObject> Prompt(string caption, string message, Collection<FieldDescription> descriptions) { throw new NotSupportedException(); }
        public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName) { throw new NotSupportedException(); }
        public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName, PSCredentialTypes allowedCredentialTypes, PSCredentialUIOptions options) { throw new NotSupportedException(); }
        public override int PromptForChoice(string caption, string message, Collection<ChoiceDescription> choices, int defaultChoice) { throw new NotSupportedException(); }
    }
    public sealed class RecordingConsoleHost : PSHost {
        private readonly Guid id = Guid.NewGuid();
        public readonly RecordingConsoleUI Recorder = new RecordingConsoleUI();
        public override Guid InstanceId { get { return id; } }
        public override string Name { get { return "WinPDFMerge recording host"; } }
        public override Version Version { get { return new Version(1,0); } }
        public override PSHostUserInterface UI { get { return Recorder; } }
        public override CultureInfo CurrentCulture { get { return CultureInfo.InvariantCulture; } }
        public override CultureInfo CurrentUICulture { get { return CultureInfo.InvariantCulture; } }
        public override void SetShouldExit(int exitCode) { throw new NotSupportedException(); }
        public override void EnterNestedPrompt() { throw new NotSupportedException(); }
        public override void ExitNestedPrompt() { throw new NotSupportedException(); }
        public override void NotifyBeginApplication() { }
        public override void NotifyEndApplication() { }
    }
}
