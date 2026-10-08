"""Read-only T19 source research record; no PDF, native app, or test execution."""
import argparse
import datetime as dt
import hashlib
import json
import pathlib
import subprocess
import sys

BASELINE = "7e0fe6fde769911f323bd87e4c4f9382d26e331b"
MAIN = "802b13da6785cc12c43158d1f6b63e6fb40a12da"
PYTHON_SHA = "dd5f8d19f6755d6491ee7c4bef2fe35ddd521334cc3ca3ed8fc93ebcadf135d0"
FILES = [
    "AGENTS.md", "docs/codex/INDEX.md", "docs/codex/STATUS.md",
    "docs/codex/NEXT_SESSION.md", "docs/codex/TASKS.json",
    "docs/codex/tasks/T19.md", "docs/codex/PRODUCT_SPEC.md",
    "docs/codex/TECHNICAL_SPEC.md", "docs/codex/TEST_STRATEGY.md",
    "docs/codex/GITHUB_WORKFLOW.md", "docs/codex/ACCEPTANCE_CASES.json",
    "WinPDFMerge.ps1", "src/WinPDFMerge.Helpers.ps1", "WinPDFMerge.bat",
    "README.md", "docs/EMAIL_PRESETS.md",
]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def git(repo, *args):
    return subprocess.check_output(["git", *args], cwd=repo, text=True).strip()


def source(title, url, quotes, observations, limits):
    quote_words = sum(len(q.split()) for q in quotes)
    assert quote_words <= 25, (url, quote_words)
    return dict(Title=title, URL=url, ExactQuotes=quotes, TotalQuotedWords=quote_words,
                Observations=observations, Limits=limits)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    repo = pathlib.Path(__file__).resolve().parents[2]
    out = pathlib.Path(args.output).resolve()
    assert out.is_relative_to(repo / "tests" / ".work")
    assert not out.exists(), "Refusing to overwrite a retained research record"
    assert sha(pathlib.Path(sys.executable)) == PYTHON_SHA, "Approved Python pin changed"
    head = git(repo, "rev-parse", "HEAD")
    assert head == BASELINE, "Research baseline changed; inspect before reusing findings"
    records = [
        source("PDFtk Server Manual", "https://www.pdflabs.com/docs/pdftk-man-page/", [
            "Use the compress filter to restore page stream compression.",
            "any XFA data in the input is automatically omitted."], [
            "cat assembles input pages into a new PDF; whole-input concatenation documents bookmark merging.",
            "compress concerns PDF page stream compression; it is not a universal feature-preservation certification.",
            "Multi-input assembly explicitly omits XFA. dump_data_annots documents only link annotations.",
            "flatten is a separate explicit output operation, absent from the current application vector."], [
            "This is an unversioned vendor manual; native behavior of the selected PDFtk Server 2.02 still requires observations.",
            "It does not establish universal AcroForm, attachment, tag, signature or archival retention."]),
        source("High Level Devices, Ghostscript 10.08.0", "https://ghostscript.readthedocs.io/en/gs10.08.0/VectorDevices.html", [
            "Instead, a new PDF file is being created"], [
            "pdfwrite recreates a PDF page description; a similar appearance is distinct from identical document structure.",
            "The manual warns that some non-marking information does not carry through.",
            "PreserveAnnots describes attempts to retain most annotations and names Link/Widget exceptions; PreserveMarkedContent describes an attempt, not accessibility validation.",
            "Outlines ordinarily pass through pdfmarks. Documentation of features is not a guarantee for this selected vector/corpus."], [
            "No signature, XFA or attachment preservation guarantee was found in the reviewed versioned page.",
            "Do not replace actual feature outcomes with blanket link/widget loss claims; interpreter behavior must be observed."]),
        source("Merging PDF files, pypdf 6.10.0", "https://pypdf.readthedocs.io/en/6.10.0/user/merging-pdfs.html", [
            "When merging forms, some form fields may have the same names, preventing access to some data."], [
            "Repeated source form names require explicit test coverage with distinct values.",
            "pypdf's top-name grouping is a library option; silently renaming/repairing application inputs would hide the PDFtk characterization."], [
            "pypdf is a development fixture/inspection dependency, not the application's merger; its merge outcomes do not prove PDFtk outcomes."]),
        source("Interactions with PDF Forms, pypdf 6.10.0", "https://pypdf.readthedocs.io/en/6.10.0/user/forms.html", [
            "Forms have a dual nature"], [
            "Catalog AcroForm/Fields and page Widget annotations are distinct views that must both be traversed.",
            "Field objects may be fused widgets or hierarchical parents with Kids. Repeated widgets for one field differ from colliding field names in separate documents.",
            "Values, parent-child associations and appearance streams need separate recording; visible form text does not establish editable-field retention."], [
            "A field-name dictionary alone can hide repeated-name collisions or orphan page widgets; no interactive GUI editing was performed here."]),
        source("PdfWriter class, pypdf 6.10.0", "https://pypdf.readthedocs.io/en/6.10.0/modules/PdfWriter.html", [
            "the original document is written first and new/modified content is appended."], [
            "The library distinguishes incremental append from a new document rewrite and discusses signed documents in that context.",
            "This application uses PDFtk cat and Ghostscript pdfwrite, not pypdf incremental writes."], [
            "This API description does not validate any signature or support a claim that arbitrary changes preserve signature permissions."]),
        source("Permissions and limitations of signed PDFs, Adobe Acrobat", "https://helpx.adobe.com/acrobat/desktop/e-sign-documents/learn-about-signatures/signed-pdf-limitations.html", [
            "Save a backup copy of the unsigned PDF before signing."], [
            "Adobe describes editing restrictions and obtaining original/unsigned copies of signed documents.",
            "Keep originals; this merger neither verifies nor promises cryptographic signature validity."], [
            "This is current Acrobat guidance, not an engine-specific PDFtk/Ghostscript acceptance result.",
            "Visible signature appearance, page readability and positive page count cannot establish cryptographic validity; no signed sample or trust-chain validation was run."]),
    ]
    report = dict(
        SchemaVersion=1, Task="T19", Result="pass", EvidenceClass="primary-source design review",
        ObservedAtUtc=dt.datetime.now(dt.timezone.utc).isoformat(), CommitUnderResearch=head,
        DirtyWorktree=bool(git(repo, "status", "--porcelain=v1")),
        CurrentStatus=git(repo, "status", "--porcelain=v1").splitlines(),
        InitialRepositoryCheck=dict(Commit=BASELINE, Branch="codex/v1.0.0-readiness", Clean=True,
                                    LiveFeature=BASELINE, LiveMain=MAIN,
                                    Origin="https://github.com/PikkuJanne/WinPDFMerger.git",
                                    Method="Actual git reads and live ls-remote in the initial T19 tool invocation; this receipt does not repeat the network query"),
        ResearchMethod="Actual web.run search/open/find of these primary vendor URLs; observations manually transcribed by independent reviewer, not fetched by this producer",
        Sources=records,
        CurrentFiles=[dict(Path=p, SHA256=sha(repo/p), Bytes=(repo/p).stat().st_size) for p in FILES],
        Producer=dict(Path="tests/.work/Research-T19Preservation.py", SHA256=sha(pathlib.Path(__file__)),
                      PythonLabel="approved bundled Python 3.12.14", PythonSHA256=PYTHON_SHA),
        CurrentNativeVectors=dict(
            Pdftk="input operands; cat; output; owned staged master; compress; dont_ask",
            Ghostscript="BATCH; NOPAUSE; SAFER; PDFSTOPONERROR; pdfwrite; CompatibilityLevel=1.6; fixed screen/ebook; DetectDuplicateImages=true; -o owned staged candidate; -f published master",
            Note="Read-only current source inspection; no native invocation, publication, runtime parameter or feature-repair change by reviewer"),
        CorpusRecommendations=[
            "Resolve nested bookmark and internal GoTo targets to output page offsets; record URI strings without visiting them.",
            "Traverse raw AcroForm leaves and page widgets including distinct repeated names/values, Parent/Kids/P and appearances; retain orphan-widget observations separately.",
            "Hash decoded attachment payloads and distinguish document EmbeddedFiles from page FileAttachment annotations.",
            "Record annotation subtype/rect/content/appearance separately from rendering; drawing an appearance is not editable annotation retention.",
            "Record raw rotation and effective page orientation separately; email may bake rotation into page content.",
            "Record minimal tags, Lang, marked-content IDs and parent-tree relationships as structural observations; do not infer accessible reading order or PDF/UA compliance.",
            "Signed-sample exclusion/non-guarantee is expressly allowed by TEST_STRATEGY; XFA/signature validity must remain excluded from proven guarantees.",
            "Observe originals, master, screen and ebook separately. If a valid candidate is larger, label its controlled retained characterization path and do not force final publication.",
            "Preserve original/foreign hashes, exact selected engine pins/vectors and clearly distinguish controlled inspection from application/native/manual acceptance."],
        SuggestedDocumentation="The master is assembled by PDFtk without intentional page rasterization or image downsampling. It is a new merged document. Feature behavior depends on the input and engine version; keep originals and check forms, links, bookmarks and attachments. The optional screen/ebook PDF is a rewrite and can change visible detail and document structure. XFA and signature validity are not supported guarantees. Neither output establishes PDF/A, accessibility certification, archival certification or malware removal.",
        BlockingFindings=[],
        AcceptanceStatus=dict(AC044="not_run", AC045="not_run"),
        Limits=["Research/design review only; no PDF authoring, application/native/suite run or actual output inspection.",
                "No visual, physical Explorer, form-editing, accessibility, signature, malware or archival certification claim.",
                "Root and fixture/test authors own implementations; later independent actual receipt/docs review remains pending.",
                "Repository-relative labels only in this receipt; no private PDF/name/profile values or downloaded third-party pages exported.",
                "No historical evidence recursion. Quotations are at most 25 words per primary webpage source."])
    out.write_text(json.dumps(report, indent=2, ensure_ascii=False)+"\n", encoding="utf-8")
    print(json.dumps(dict(Result="pass", Report=out.relative_to(repo).as_posix(),
                          ReportSHA256=sha(out), Sources=len(records), Acceptance="not_run")))


if __name__ == "__main__":
    main()
