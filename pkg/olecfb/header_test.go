package olecfb

import (
	"encoding/binary"
	"io"
	"reflect"
	"strings"
	"testing"

	"github.com/KairuiKin/go-olespec/pkg/oleps"
)

func TestOpenBytes_ValidV3Header(t *testing.T) {
	buf := buildValidHeader(cfbMajorVersion3)
	f, err := OpenBytes(buf, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	if f == nil {
		t.Fatal("OpenBytes returned nil file")
	}
	if f.hdr == nil {
		t.Fatal("file header was not parsed")
	}
	if f.hdr.MajorVersion != cfbMajorVersion3 {
		t.Fatalf("unexpected major version: %d", f.hdr.MajorVersion)
	}
	if f.root.Name != "Root Entry" {
		t.Fatalf("unexpected root name: %s", f.root.Name)
	}
}

func TestOpenBytes_ParseRootDirectoryEntry(t *testing.T) {
	buf := buildValidFileWithRootEntry("My Root")
	f, err := OpenBytes(buf, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	if f.root.Name != "My Root" {
		t.Fatalf("unexpected root name: %q", f.root.Name)
	}
	if f.root.StateBits != 0x11223344 {
		t.Fatalf("unexpected root state bits: 0x%08X", f.root.StateBits)
	}
}

func TestOpenBytes_InvalidSignature(t *testing.T) {
	buf := buildValidHeader(cfbMajorVersion3)
	buf[0] = 0
	_, err := OpenBytes(buf, OpenOptions{Mode: Strict})
	if err == nil {
		t.Fatal("expected error, got nil")
	}
	if !IsCode(err, ErrBadHeader) {
		t.Fatalf("expected ErrBadHeader, got %v", err)
	}
}

func TestOpenBytes_TooSmall(t *testing.T) {
	_, err := OpenBytes(make([]byte, 100), OpenOptions{Mode: Strict})
	if err == nil {
		t.Fatal("expected error, got nil")
	}
	if !IsCode(err, ErrBadHeader) {
		t.Fatalf("expected ErrBadHeader, got %v", err)
	}
}

func TestCreateInMemory_BeginTransaction(t *testing.T) {
	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if tx == nil {
		t.Fatal("Begin returned nil transaction")
	}
}

func TestOpenBytes_LenientMode_LoadFATWarning(t *testing.T) {
	buf := buildValidHeader(cfbMajorVersion3)
	binary.LittleEndian.PutUint32(buf[44:48], 1) // NumFATSectors
	binary.LittleEndian.PutUint32(buf[68:72], 0) // FirstDIFAT not end-of-chain => extended DIFAT path
	binary.LittleEndian.PutUint32(buf[72:76], 1) // NumDIFATSectors

	_, strictErr := OpenBytes(buf, OpenOptions{Mode: Strict})
	if strictErr == nil {
		t.Fatal("strict mode should fail for unsupported extended DIFAT")
	}

	f, err := OpenBytes(buf, OpenOptions{Mode: Lenient})
	if err != nil {
		t.Fatalf("lenient mode should not fail: %v", err)
	}
	rep := f.Report()
	if !rep.Partial {
		t.Fatal("expected partial report in lenient mode")
	}
	if len(rep.Warnings) == 0 {
		t.Fatal("expected warnings in lenient mode")
	}
}

func TestOpenBytes_LenientMode_RootOutOfBounds(t *testing.T) {
	buf := buildValidHeader(cfbMajorVersion3)
	binary.LittleEndian.PutUint32(buf[48:52], 99) // first directory sector out of bounds

	_, strictErr := OpenBytes(buf, OpenOptions{Mode: Strict})
	if strictErr == nil {
		t.Fatal("strict mode should fail for out-of-bounds directory")
	}

	f, err := OpenBytes(buf, OpenOptions{Mode: Lenient})
	if err != nil {
		t.Fatalf("lenient mode should not fail: %v", err)
	}
	if f.root.Name != "Root Entry" {
		t.Fatalf("unexpected fallback root name: %s", f.root.Name)
	}
	rep := f.Report()
	if !rep.Partial || len(rep.Warnings) == 0 {
		t.Fatal("expected partial report with warning")
	}
}

func TestOpenBytes_LoadFATFromDifatHeader(t *testing.T) {
	buf := buildValidFileWithFatAndRoot("Root Entry")
	f, err := OpenBytes(buf, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	if got, want := len(f.fat), cfbHeaderSize/4; got != want {
		t.Fatalf("unexpected fat entry count: got %d want %d", got, want)
	}
	if f.fat[0] != cfbFatSector {
		t.Fatalf("unexpected FAT[0]: got 0x%08X", f.fat[0])
	}
	if f.fat[1] != cfbEndOfChain {
		t.Fatalf("unexpected FAT[1]: got 0x%08X", f.fat[1])
	}
}

func TestOpenBytes_BuildDirectoryTreeAndWalk(t *testing.T) {
	buf := buildValidFileWithDirectoryTree()
	f, err := OpenBytes(buf, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	if got, want := len(f.nodes), 3; got != want {
		t.Fatalf("unexpected node count: got %d want %d", got, want)
	}

	var paths []string
	if err := f.Walk(func(n Node) error {
		paths = append(paths, n.Path)
		return nil
	}); err != nil {
		t.Fatalf("Walk returned error: %v", err)
	}
	wantPaths := []string{"/", "/Folder", "/Folder/Doc"}
	if !reflect.DeepEqual(paths, wantPaths) {
		t.Fatalf("unexpected walk paths: got %#v want %#v", paths, wantPaths)
	}

	folder := f.nodes[1]
	if folder.Type != NodeStorage {
		t.Fatalf("unexpected folder type: %v", folder.Type)
	}
	if folder.ChildCount != 1 {
		t.Fatalf("unexpected folder child count: %d", folder.ChildCount)
	}
	stream := f.nodes[2]
	if stream.Type != NodeStream {
		t.Fatalf("unexpected stream type: %v", stream.Type)
	}
	if stream.ParentID != 1 {
		t.Fatalf("unexpected stream parent: %d", stream.ParentID)
	}
}

func TestOpenStream_NormalFATStream_V4(t *testing.T) {
	fileBytes, payload := buildValidV4FileWithSingleNormalStream()
	f, err := OpenBytes(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	sr, err := f.OpenStream("/Blob")
	if err != nil {
		t.Fatalf("OpenStream returned error: %v", err)
	}
	defer sr.Close()

	got := make([]byte, len(payload))
	if _, err := io.ReadFull(sr, got); err != nil {
		t.Fatalf("ReadFull returned error: %v", err)
	}
	if !reflect.DeepEqual(got, payload) {
		t.Fatalf("unexpected stream payload")
	}
}

func TestOpenPropertySet_SummaryInformation(t *testing.T) {
	ps := buildSummaryPropertySetStreamBytes("Core Title", 9)
	fileBytes, _ := buildValidV4FileWithNamedStream("\x05SummaryInformation", ps)
	f, err := OpenBytes(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	set, err := f.OpenSummaryInformation()
	if err != nil {
		t.Fatalf("OpenSummaryInformation returned error: %v", err)
	}
	title, ok := set.GetString(oleps.PIDTitle)
	if !ok {
		t.Fatal("title property not found")
	}
	if title != "Core Title" {
		t.Fatalf("unexpected title: %q", title)
	}
	pages, ok := set.GetInt64(oleps.PIDPageCount)
	if !ok || pages != 9 {
		t.Fatalf("unexpected page count: %d", pages)
	}
}

func TestOpenPropertySet_DocumentSummaryInformation(t *testing.T) {
	ps := buildDocumentSummaryPropertySetStreamBytes("Doc Author")
	fileBytes, _ := buildValidV4FileWithNamedStream("\x05DocumentSummaryInformation", ps)
	f, err := OpenBytes(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	set, err := f.OpenDocumentSummaryInformation()
	if err != nil {
		t.Fatalf("OpenDocumentSummaryInformation returned error: %v", err)
	}
	author, ok := set.GetString(oleps.PIDAuthor)
	if !ok {
		t.Fatal("author property not found")
	}
	if author != "Doc Author" {
		t.Fatalf("unexpected author: %q", author)
	}
}

func TestExtract_FromV4Stream(t *testing.T) {
	fileBytes, payload := buildValidV4FileWithSingleNormalStream()
	f, err := OpenBytes(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	rep, err := f.Extract(ExtractOptions{Deduplicate: true})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}
	if rep.Stats.ArtifactsTotal != 1 {
		t.Fatalf("unexpected artifact total: %d", rep.Stats.ArtifactsTotal)
	}
	a := rep.Artifacts[0]
	if a.Path != "/Blob" {
		t.Fatalf("unexpected artifact path: %s", a.Path)
	}
	if a.Kind != ArtifactStream {
		t.Fatalf("unexpected artifact kind: %s", a.Kind)
	}
	if a.Status != ArtifactOK {
		t.Fatalf("unexpected artifact status: %s", a.Status)
	}
	if a.Size != int64(len(payload)) {
		t.Fatalf("unexpected artifact size: %d", a.Size)
	}
	if a.SHA256 == "" {
		t.Fatal("expected sha256")
	}
}

func TestExtract_DetectOLEDS(t *testing.T) {
	payload := buildOle10NativeBytes("a.txt", "C:\\a.txt", []byte("abc"))
	fileBytes, _ := buildValidV4FileWithNamedStream("\x01Ole10Native", payload)
	f, err := OpenBytes(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	rep, err := f.Extract(ExtractOptions{Deduplicate: true, DetectOLEDS: true})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}
	if len(rep.Artifacts) != 1 {
		t.Fatalf("unexpected artifact count: %d", len(rep.Artifacts))
	}
	a := rep.Artifacts[0]
	if a.Kind != ArtifactOleObj {
		t.Fatalf("expected ole object kind, got %s", a.Kind)
	}
	if a.Note != "oleds:ole10native" {
		t.Fatalf("unexpected oleds note: %q", a.Note)
	}
	if a.OLEFileName != "a.txt" {
		t.Fatalf("unexpected oleds file name: %q", a.OLEFileName)
	}
	if a.OLESourcePath != "C:\\a.txt" {
		t.Fatalf("unexpected oleds source path: %q", a.OLESourcePath)
	}
}

func TestExtract_UnwrapOle10Native(t *testing.T) {
	innerBytes, _ := buildValidV4FileWithSingleNormalStream()
	payload := buildOle10NativeBytes("inner.cfb", "C:\\inner.cfb", innerBytes)
	fileBytes := buildValidV4FileWithBigNamedStream("\x01Ole10Native", payload)
	f, err := OpenBytes(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	rep, err := f.Extract(ExtractOptions{
		Deduplicate:       false,
		DetectOLEDS:       true,
		UnwrapOle10Native: true,
		Limits:            ExtractLimits{MaxDepth: 4},
	})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}

	var parent, child, grand *Artifact
	for i := range rep.Artifacts {
		a := &rep.Artifacts[i]
		switch a.Path {
		case "/\x01Ole10Native":
			parent = a
		case "/\x01Ole10Native!$ole10native":
			child = a
		case "/\x01Ole10Native!$ole10native!/Blob":
			grand = a
		}
	}
	if parent == nil {
		t.Fatal("parent Ole10Native stream artifact not found")
	}
	if child == nil {
		t.Fatal("unwrapped Ole10Native payload artifact not found")
	}
	if grand == nil {
		t.Fatal("nested OLE child artifact not found")
	}
	if child.ParentID != parent.ID {
		t.Fatalf("unexpected unwrapped parent id: got %q want %q", child.ParentID, parent.ID)
	}
	if grand.ParentID != child.ID {
		t.Fatalf("unexpected nested parent id: got %q want %q", grand.ParentID, child.ID)
	}
	if !strings.HasPrefix(child.Note, "ole10native;file=inner.cfb") {
		t.Fatalf("unexpected unwrapped note: %q", child.Note)
	}
	if child.OLEFileName != "inner.cfb" {
		t.Fatalf("unexpected unwrapped file name: %q", child.OLEFileName)
	}
	if child.OLESourcePath != "C:\\inner.cfb" {
		t.Fatalf("unexpected unwrapped source path: %q", child.OLESourcePath)
	}
}

func TestExtract_UnwrapOle10NativeDisabled(t *testing.T) {
	innerBytes, _ := buildValidV4FileWithSingleNormalStream()
	payload := buildOle10NativeBytes("inner.cfb", "C:\\inner.cfb", innerBytes)
	fileBytes := buildValidV4FileWithBigNamedStream("\x01Ole10Native", payload)
	f, err := OpenBytes(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	rep, err := f.Extract(ExtractOptions{
		Deduplicate:       false,
		DetectOLEDS:       true,
		UnwrapOle10Native: false,
	})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}
	for _, a := range rep.Artifacts {
		if a.Path == "/\x01Ole10Native!$ole10native" || a.Path == "/\x01Ole10Native!$ole10native!/Blob" {
			t.Fatalf("unexpected unwrapped artifact when disabled: %s", a.Path)
		}
	}
}

func TestExtract_DetectImages(t *testing.T) {
	pngPrefix := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A}
	fileBytes, _ := buildValidV4FileWithNamedStream("Image1", pngPrefix)
	f, err := OpenBytes(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	rep, err := f.Extract(ExtractOptions{DetectImages: true})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}
	if len(rep.Artifacts) != 1 {
		t.Fatalf("unexpected artifact count: %d", len(rep.Artifacts))
	}
	a := rep.Artifacts[0]
	if a.Kind != ArtifactImage {
		t.Fatalf("expected image artifact, got %s", a.Kind)
	}
	if a.MediaType != "image/png" {
		t.Fatalf("unexpected media type: %q", a.MediaType)
	}
}

func TestExtract_IncludeRaw(t *testing.T) {
	fileBytes, payload := buildValidV4FileWithSingleNormalStream()
	f, err := OpenBytes(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}

	rep, err := f.Extract(ExtractOptions{IncludeRaw: false})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}
	if len(rep.Artifacts) != 1 {
		t.Fatalf("unexpected artifact count: %d", len(rep.Artifacts))
	}
	if rep.Artifacts[0].Raw != nil {
		t.Fatal("raw bytes should be nil when IncludeRaw=false")
	}

	rep2, err := f.Extract(ExtractOptions{IncludeRaw: true})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}
	if len(rep2.Artifacts) != 1 {
		t.Fatalf("unexpected artifact count: %d", len(rep2.Artifacts))
	}
	raw := rep2.Artifacts[0].Raw
	if len(raw) != len(payload) {
		t.Fatalf("unexpected raw length: got %d want %d", len(raw), len(payload))
	}
	if !reflect.DeepEqual(raw, payload) {
		t.Fatal("unexpected raw payload")
	}
}

func TestExtract_RecursiveNestedOLE(t *testing.T) {
	innerBytes, _ := buildValidV4FileWithSingleNormalStream()
	outerBytes := buildValidV4FileWithBigNamedStream("Embedded", innerBytes)

	f, err := OpenBytes(outerBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	rep, err := f.Extract(ExtractOptions{Deduplicate: false, Limits: ExtractLimits{MaxDepth: 4}})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}

	var outer, child *Artifact
	for i := range rep.Artifacts {
		a := &rep.Artifacts[i]
		if a.Path == "/Embedded" {
			outer = a
		}
		if a.Path == "/Embedded!/Blob" {
			child = a
		}
	}
	if outer == nil {
		t.Fatal("outer embedded artifact not found")
	}
	if child == nil {
		t.Fatal("nested child artifact not found")
	}
	if outer.Kind != ArtifactOLEFile {
		t.Fatalf("unexpected outer artifact kind: %s", outer.Kind)
	}
	if child.ParentID != outer.ID {
		t.Fatalf("unexpected child parent id: got %q want %q", child.ParentID, outer.ID)
	}
	if outer.Children == 0 {
		t.Fatal("outer artifact should have nested children")
	}
}

func TestExtract_RecursiveDepthLimit(t *testing.T) {
	innerBytes, _ := buildValidV4FileWithSingleNormalStream()
	midBytes := buildValidV4FileWithBigNamedStream("InnerOLE", innerBytes)
	outerBytes := buildValidV4FileWithBigNamedStream("Embedded", midBytes)

	f, err := OpenBytes(outerBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	rep, err := f.Extract(ExtractOptions{Deduplicate: false, Limits: ExtractLimits{MaxDepth: 1}})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}
	foundMid := false
	foundLeaf := false
	foundDepthWarning := false
	for _, a := range rep.Artifacts {
		if a.Path == "/Embedded!/InnerOLE" {
			foundMid = true
		}
		if a.Path == "/Embedded!/InnerOLE!/Blob" {
			foundLeaf = true
		}
	}
	for _, w := range rep.Warnings {
		if w.Code == ErrDepthExceeded {
			foundDepthWarning = true
			break
		}
	}
	if !foundMid {
		t.Fatal("expected mid-level artifact")
	}
	if foundLeaf {
		t.Fatal("leaf artifact should not be extracted when max depth is 1")
	}
	if !foundDepthWarning {
		t.Fatal("expected depth exceeded warning")
	}
}

func TestOpenStream_MiniStream(t *testing.T) {
	buf, payload := buildValidFileWithMiniStream()
	f, err := OpenBytes(buf, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	sr, err := f.OpenStream("/Small")
	if err == nil {
		defer sr.Close()
	}
	if err != nil {
		t.Fatalf("OpenStream returned error: %v", err)
	}
	got := make([]byte, len(payload))
	if _, err := io.ReadFull(sr, got); err != nil {
		t.Fatalf("ReadFull returned error: %v", err)
	}
	if !reflect.DeepEqual(got, payload) {
		t.Fatalf("unexpected mini stream payload")
	}
}

func TestWalkEx_DFSAndBFSOrder(t *testing.T) {
	buf := buildValidFileWithBranchingTree()
	f, err := OpenBytes(buf, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}

	var dfs []string
	_, err = f.WalkEx(WalkOptions{IncludeRoot: true, Order: WalkDFS}, func(ev WalkEvent) error {
		dfs = append(dfs, ev.Node.Path)
		return nil
	})
	if err != nil {
		t.Fatalf("WalkEx DFS returned error: %v", err)
	}
	wantDFS := []string{"/", "/A", "/A/A1", "/B"}
	if !reflect.DeepEqual(dfs, wantDFS) {
		t.Fatalf("unexpected DFS order: got %#v want %#v", dfs, wantDFS)
	}

	var bfs []string
	_, err = f.WalkEx(WalkOptions{IncludeRoot: true, Order: WalkBFS}, func(ev WalkEvent) error {
		bfs = append(bfs, ev.Node.Path)
		return nil
	})
	if err != nil {
		t.Fatalf("WalkEx BFS returned error: %v", err)
	}
	wantBFS := []string{"/", "/A", "/B", "/A/A1"}
	if !reflect.DeepEqual(bfs, wantBFS) {
		t.Fatalf("unexpected BFS order: got %#v want %#v", bfs, wantBFS)
	}
}

func TestGetNodeAPIs(t *testing.T) {
	buf := buildValidFileWithBranchingTree()
	f, err := OpenBytes(buf, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	n, err := f.GetNodeByPath("/A/A1")
	if err != nil {
		t.Fatalf("GetNodeByPath returned error: %v", err)
	}
	if n.Type != NodeStream {
		t.Fatalf("unexpected node type: %v", n.Type)
	}
	n2, err := f.GetNode(n.ID)
	if err != nil {
		t.Fatalf("GetNode returned error: %v", err)
	}
	if n2.Path != "/A/A1" {
		t.Fatalf("unexpected node path: %s", n2.Path)
	}
	if got := len(f.ListNodes()); got != 4 {
		t.Fatalf("unexpected node list count: %d", got)
	}
}

func TestOpenBytes_MaxObjectCountLimit(t *testing.T) {
	buf := buildValidFileWithBranchingTree()

	_, strictErr := OpenBytes(buf, OpenOptions{Mode: Strict, MaxObjectCount: 2})
	if strictErr == nil {
		t.Fatal("strict mode should fail on object count limit")
	}
	if !IsCode(strictErr, ErrLimitExceeded) {
		t.Fatalf("expected ErrLimitExceeded, got %v", strictErr)
	}

	f, err := OpenBytes(buf, OpenOptions{Mode: Lenient, MaxObjectCount: 2})
	if err != nil {
		t.Fatalf("lenient mode should not fail: %v", err)
	}
	rep := f.Report()
	if !rep.Partial || len(rep.Warnings) == 0 {
		t.Fatal("expected partial report with warning")
	}
}

func TestOpenBytes_MaxTotalBytesLimit(t *testing.T) {
	buf := buildValidFileWithBranchingTree()
	_, err := OpenBytes(buf, OpenOptions{
		Mode:          Strict,
		MaxTotalBytes: int64(len(buf) - 1),
	})
	if err == nil {
		t.Fatal("expected quota exceeded error")
	}
	if !IsCode(err, ErrQuotaExceeded) {
		t.Fatalf("expected ErrQuotaExceeded, got %v", err)
	}
}

func buildValidHeader(major uint16) []byte {
	buf := make([]byte, cfbHeaderSize)
	copy(buf[0:8], cfbSignature[:])
	// CLSID remains zero.
	binary.LittleEndian.PutUint16(buf[24:26], 0x003E) // minor version
	binary.LittleEndian.PutUint16(buf[26:28], major)
	binary.LittleEndian.PutUint16(buf[28:30], cfbByteOrder)
	if major == cfbMajorVersion4 {
		binary.LittleEndian.PutUint16(buf[30:32], cfbSectorShiftV4)
	} else {
		binary.LittleEndian.PutUint16(buf[30:32], cfbSectorShiftV3)
	}
	binary.LittleEndian.PutUint16(buf[32:34], cfbMiniSectorShift)
	// reserved [34:40] remains zero
	if major == cfbMajorVersion3 {
		binary.LittleEndian.PutUint32(buf[40:44], 0)
	}
	binary.LittleEndian.PutUint32(buf[44:48], 0)                   // FAT sectors
	binary.LittleEndian.PutUint32(buf[48:52], cfbEndOfChain)       // first dir sector
	binary.LittleEndian.PutUint32(buf[52:56], 0)                   // transaction signature
	binary.LittleEndian.PutUint32(buf[56:60], cfbMiniStreamCutoff) // mini stream cutoff
	binary.LittleEndian.PutUint32(buf[60:64], cfbEndOfChain)
	binary.LittleEndian.PutUint32(buf[64:68], 0)
	binary.LittleEndian.PutUint32(buf[68:72], cfbEndOfChain)
	binary.LittleEndian.PutUint32(buf[72:76], 0)
	for i := 0; i < cfbNumDifatEntries; i++ {
		start := 76 + i*4
		binary.LittleEndian.PutUint32(buf[start:start+4], cfbFreeSector)
	}
	return buf
}

func buildValidFileWithRootEntry(rootName string) []byte {
	header := buildValidHeader(cfbMajorVersion3)
	binary.LittleEndian.PutUint32(header[48:52], 0) // first directory sector index
	sector := make([]byte, cfbHeaderSize)

	encoded := utf16LE(rootName)
	nameLen := len(encoded) + 2
	if nameLen > 64 {
		nameLen = 64
	}
	copy(sector[0:nameLen-2], encoded[:nameLen-2])
	sector[nameLen-2] = 0
	sector[nameLen-1] = 0
	binary.LittleEndian.PutUint16(sector[64:66], uint16(nameLen))
	sector[66] = 5 // root storage
	binary.LittleEndian.PutUint32(sector[68:72], cfbNoStream)
	binary.LittleEndian.PutUint32(sector[72:76], cfbNoStream)
	binary.LittleEndian.PutUint32(sector[76:80], cfbNoStream)
	binary.LittleEndian.PutUint32(sector[96:100], 0x11223344)
	binary.LittleEndian.PutUint64(sector[120:128], 4096)

	return append(header, sector...)
}

func buildValidFileWithFatAndRoot(rootName string) []byte {
	header := buildValidHeader(cfbMajorVersion3)
	binary.LittleEndian.PutUint32(header[44:48], 1) // one FAT sector
	binary.LittleEndian.PutUint32(header[48:52], 1) // first directory sector
	binary.LittleEndian.PutUint32(header[76:80], 0) // DIFAT[0] -> FAT sector id 0

	fatSector := make([]byte, cfbHeaderSize)
	// Sector 0 is FAT sector.
	binary.LittleEndian.PutUint32(fatSector[0:4], cfbFatSector)
	// Sector 1 is directory sector and ends immediately.
	binary.LittleEndian.PutUint32(fatSector[4:8], cfbEndOfChain)
	for i := 2; i < cfbHeaderSize/4; i++ {
		start := i * 4
		binary.LittleEndian.PutUint32(fatSector[start:start+4], cfbFreeSector)
	}

	dirSector := make([]byte, cfbHeaderSize)
	encoded := utf16LE(rootName)
	nameLen := len(encoded) + 2
	if nameLen > 64 {
		nameLen = 64
	}
	copy(dirSector[0:nameLen-2], encoded[:nameLen-2])
	dirSector[nameLen-2] = 0
	dirSector[nameLen-1] = 0
	binary.LittleEndian.PutUint16(dirSector[64:66], uint16(nameLen))
	dirSector[66] = 5 // root storage
	binary.LittleEndian.PutUint32(dirSector[68:72], cfbNoStream)
	binary.LittleEndian.PutUint32(dirSector[72:76], cfbNoStream)
	binary.LittleEndian.PutUint32(dirSector[76:80], cfbNoStream)

	out := append(header, fatSector...)
	return append(out, dirSector...)
}

func buildValidFileWithDirectoryTree() []byte {
	header := buildValidHeader(cfbMajorVersion3)
	binary.LittleEndian.PutUint32(header[44:48], 1) // one FAT sector
	binary.LittleEndian.PutUint32(header[48:52], 1) // first directory sector
	binary.LittleEndian.PutUint32(header[76:80], 0) // DIFAT[0] -> FAT sector id 0

	fatSector := make([]byte, cfbHeaderSize)
	// sector 0: FAT
	binary.LittleEndian.PutUint32(fatSector[0:4], cfbFatSector)
	// sector 1: directory
	binary.LittleEndian.PutUint32(fatSector[4:8], cfbEndOfChain)
	for i := 2; i < cfbHeaderSize/4; i++ {
		start := i * 4
		binary.LittleEndian.PutUint32(fatSector[start:start+4], cfbFreeSector)
	}

	dirSector := make([]byte, cfbHeaderSize)
	writeDirEntry(dirSector[0:128], "Root Entry", 5, cfbNoStream, cfbNoStream, 1, cfbEndOfChain, 0)
	writeDirEntry(dirSector[128:256], "Folder", 1, cfbNoStream, cfbNoStream, 2, cfbEndOfChain, 0)
	writeDirEntry(dirSector[256:384], "Doc", 2, cfbNoStream, cfbNoStream, cfbNoStream, cfbEndOfChain, 11)
	// remaining entries left zero as unallocated.

	out := append(header, fatSector...)
	return append(out, dirSector...)
}

func buildValidV4FileWithSingleNormalStream() ([]byte, []byte) {
	return buildValidV4FileWithNamedStream("Blob", []byte("go-olespec-normal-stream"))
}

func buildValidV4FileWithNamedStream(streamName string, payloadPrefix []byte) ([]byte, []byte) {
	const sectorSize = 4096
	header := buildValidHeader(cfbMajorVersion4)
	binary.LittleEndian.PutUint32(header[40:44], 1) // directory sector count
	binary.LittleEndian.PutUint32(header[44:48], 1) // FAT sector count
	binary.LittleEndian.PutUint32(header[48:52], 1) // first directory sector
	binary.LittleEndian.PutUint32(header[76:80], 0) // DIFAT[0] = sector 0 (FAT)

	fatSector := make([]byte, sectorSize)
	binary.LittleEndian.PutUint32(fatSector[0:4], cfbFatSector)   // sector 0 is FAT
	binary.LittleEndian.PutUint32(fatSector[4:8], cfbEndOfChain)  // sector 1 is directory
	binary.LittleEndian.PutUint32(fatSector[8:12], cfbEndOfChain) // sector 2 is stream
	for i := 3; i < sectorSize/4; i++ {
		start := i * 4
		binary.LittleEndian.PutUint32(fatSector[start:start+4], cfbFreeSector)
	}

	dirSector := make([]byte, sectorSize)
	writeDirEntry(dirSector[0:128], "Root Entry", 5, cfbNoStream, cfbNoStream, 1, cfbEndOfChain, 0)
	writeDirEntry(dirSector[128:256], streamName, 2, cfbNoStream, cfbNoStream, cfbNoStream, 2, uint64(sectorSize))

	payload := make([]byte, sectorSize)
	copy(payload, payloadPrefix)

	paddedHeader := make([]byte, sectorSize)
	copy(paddedHeader, header)

	out := append(paddedHeader, fatSector...)
	out = append(out, dirSector...)
	out = append(out, payload...)
	return out, payload
}

func buildValidV4FileWithBigNamedStream(streamName string, payload []byte) []byte {
	const sectorSize = 4096
	header := buildValidHeader(cfbMajorVersion4)
	streamSectors := (len(payload) + sectorSize - 1) / sectorSize
	if streamSectors == 0 {
		streamSectors = 1
	}
	binary.LittleEndian.PutUint32(header[40:44], 1) // directory sectors
	binary.LittleEndian.PutUint32(header[44:48], 1) // FAT sectors
	binary.LittleEndian.PutUint32(header[48:52], 1) // first directory sector
	binary.LittleEndian.PutUint32(header[76:80], 0) // DIFAT[0] = FAT sector

	fatSector := make([]byte, sectorSize)
	binary.LittleEndian.PutUint32(fatSector[0:4], cfbFatSector)  // sector 0: FAT
	binary.LittleEndian.PutUint32(fatSector[4:8], cfbEndOfChain) // sector 1: directory
	for i := 0; i < streamSectors; i++ {
		sid := 2 + i
		next := uint32(cfbEndOfChain)
		if i+1 < streamSectors {
			next = uint32(sid + 1)
		}
		binary.LittleEndian.PutUint32(fatSector[sid*4:sid*4+4], next)
	}
	for i := 2 + streamSectors; i < sectorSize/4; i++ {
		binary.LittleEndian.PutUint32(fatSector[i*4:i*4+4], cfbFreeSector)
	}

	dirSector := make([]byte, sectorSize)
	writeDirEntry(dirSector[0:128], "Root Entry", 5, cfbNoStream, cfbNoStream, 1, cfbEndOfChain, 0)
	writeDirEntry(dirSector[128:256], streamName, 2, cfbNoStream, cfbNoStream, cfbNoStream, 2, uint64(len(payload)))

	paddedHeader := make([]byte, sectorSize)
	copy(paddedHeader, header)

	out := append(paddedHeader, fatSector...)
	out = append(out, dirSector...)
	for i := 0; i < streamSectors; i++ {
		sec := make([]byte, sectorSize)
		start := i * sectorSize
		end := start + sectorSize
		if end > len(payload) {
			end = len(payload)
		}
		if start < len(payload) {
			copy(sec, payload[start:end])
		}
		out = append(out, sec...)
	}
	return out
}

func buildSummaryPropertySetStreamBytes(title string, pageCount int32) []byte {
	headerSize := 28 + 20
	header := make([]byte, headerSize)
	binary.LittleEndian.PutUint16(header[0:2], 0xFFFE)
	binary.LittleEndian.PutUint16(header[2:4], 0x0000)
	binary.LittleEndian.PutUint32(header[24:28], 1)
	copy(header[28:44], oleps.FMTIDSummaryInformation[:])
	binary.LittleEndian.PutUint32(header[44:48], uint32(headerSize))

	titleUTF16 := utf16LE(title + "\x00")
	valTitle := make([]byte, 8+len(titleUTF16))
	binary.LittleEndian.PutUint16(valTitle[0:2], uint16(oleps.VTLPWSTR))
	binary.LittleEndian.PutUint32(valTitle[4:8], uint32(len(title)+1))
	copy(valTitle[8:], titleUTF16)

	valPages := make([]byte, 8)
	binary.LittleEndian.PutUint16(valPages[0:2], uint16(oleps.VTI4))
	binary.LittleEndian.PutUint32(valPages[4:8], uint32(pageCount))

	offTitle := uint32(24)
	offPages := offTitle + uint32(len(valTitle))
	sectionSize := int(offPages) + len(valPages)
	section := make([]byte, sectionSize)
	binary.LittleEndian.PutUint32(section[0:4], uint32(sectionSize))
	binary.LittleEndian.PutUint32(section[4:8], 2)
	binary.LittleEndian.PutUint32(section[8:12], oleps.PIDTitle)
	binary.LittleEndian.PutUint32(section[12:16], offTitle)
	binary.LittleEndian.PutUint32(section[16:20], oleps.PIDPageCount)
	binary.LittleEndian.PutUint32(section[20:24], offPages)
	copy(section[offTitle:], valTitle)
	copy(section[offPages:], valPages)

	return append(header, section...)
}

func buildDocumentSummaryPropertySetStreamBytes(author string) []byte {
	headerSize := 28 + 20
	header := make([]byte, headerSize)
	binary.LittleEndian.PutUint16(header[0:2], 0xFFFE)
	binary.LittleEndian.PutUint16(header[2:4], 0x0000)
	binary.LittleEndian.PutUint32(header[24:28], 1)
	copy(header[28:44], oleps.FMTIDDocumentSummaryInformation[:])
	binary.LittleEndian.PutUint32(header[44:48], uint32(headerSize))

	authorUTF16 := utf16LE(author + "\x00")
	valAuthor := make([]byte, 8+len(authorUTF16))
	binary.LittleEndian.PutUint16(valAuthor[0:2], uint16(oleps.VTLPWSTR))
	binary.LittleEndian.PutUint32(valAuthor[4:8], uint32(len(author)+1))
	copy(valAuthor[8:], authorUTF16)

	offAuthor := uint32(16)
	sectionSize := int(offAuthor) + len(valAuthor)
	section := make([]byte, sectionSize)
	binary.LittleEndian.PutUint32(section[0:4], uint32(sectionSize))
	binary.LittleEndian.PutUint32(section[4:8], 1)
	binary.LittleEndian.PutUint32(section[8:12], oleps.PIDAuthor)
	binary.LittleEndian.PutUint32(section[12:16], offAuthor)
	copy(section[offAuthor:], valAuthor)

	return append(header, section...)
}

func buildOle10NativeBytes(fileName, sourcePath string, payload []byte) []byte {
	var body []byte
	body = append(body, 0x02, 0x00)
	body = append(body, []byte(fileName)...)
	body = append(body, 0)
	body = append(body, []byte(sourcePath)...)
	body = append(body, 0)
	body = append(body, []byte(sourcePath)...)
	body = append(body, 0)
	sz := make([]byte, 4)
	binary.LittleEndian.PutUint32(sz, uint32(len(payload)))
	body = append(body, sz...)
	body = append(body, payload...)

	out := make([]byte, 4)
	binary.LittleEndian.PutUint32(out, uint32(len(body)))
	out = append(out, body...)
	return out
}

func buildValidFileWithBranchingTree() []byte {
	header := buildValidHeader(cfbMajorVersion3)
	binary.LittleEndian.PutUint32(header[44:48], 1) // one FAT sector
	binary.LittleEndian.PutUint32(header[48:52], 1) // first directory sector
	binary.LittleEndian.PutUint32(header[76:80], 0) // DIFAT[0] -> FAT sector id 0

	fatSector := make([]byte, cfbHeaderSize)
	binary.LittleEndian.PutUint32(fatSector[0:4], cfbFatSector)   // sector 0 FAT
	binary.LittleEndian.PutUint32(fatSector[4:8], cfbEndOfChain)  // sector 1 directory
	binary.LittleEndian.PutUint32(fatSector[8:12], cfbEndOfChain) // sector 2 stream data
	for i := 3; i < cfbHeaderSize/4; i++ {
		start := i * 4
		binary.LittleEndian.PutUint32(fatSector[start:start+4], cfbFreeSector)
	}

	dirSector := make([]byte, cfbHeaderSize)
	// Root child points to A.
	writeDirEntry(dirSector[0:128], "Root Entry", 5, cfbNoStream, cfbNoStream, 1, cfbEndOfChain, 0)
	// A has right sibling B, and child A1.
	writeDirEntry(dirSector[128:256], "A", 1, cfbNoStream, 2, 3, cfbEndOfChain, 0)
	// B storage.
	writeDirEntry(dirSector[256:384], "B", 1, cfbNoStream, cfbNoStream, cfbNoStream, cfbEndOfChain, 0)
	// A1 stream under A.
	writeDirEntry(dirSector[384:512], "A1", 2, cfbNoStream, cfbNoStream, cfbNoStream, 2, 4096)

	// Stream data sector (unused by this test but keeps entry coherent).
	streamSector := make([]byte, cfbHeaderSize)
	copy(streamSector, []byte("A1-stream"))

	out := append(header, fatSector...)
	out = append(out, dirSector...)
	out = append(out, streamSector...)
	return out
}

func buildValidFileWithMiniStream() ([]byte, []byte) {
	header := buildValidHeader(cfbMajorVersion3)
	binary.LittleEndian.PutUint32(header[44:48], 1) // one FAT sector
	binary.LittleEndian.PutUint32(header[48:52], 1) // first directory sector
	binary.LittleEndian.PutUint32(header[60:64], 2) // first MiniFAT sector
	binary.LittleEndian.PutUint32(header[64:68], 1) // num MiniFAT sectors
	binary.LittleEndian.PutUint32(header[76:80], 0) // DIFAT[0] -> FAT sector 0

	fatSector := make([]byte, cfbHeaderSize)
	binary.LittleEndian.PutUint32(fatSector[0:4], cfbFatSector)    // sector 0 FAT
	binary.LittleEndian.PutUint32(fatSector[4:8], cfbEndOfChain)   // sector 1 directory
	binary.LittleEndian.PutUint32(fatSector[8:12], cfbEndOfChain)  // sector 2 miniFAT
	binary.LittleEndian.PutUint32(fatSector[12:16], cfbEndOfChain) // sector 3 root mini stream data
	for i := 4; i < cfbHeaderSize/4; i++ {
		start := i * 4
		binary.LittleEndian.PutUint32(fatSector[start:start+4], cfbFreeSector)
	}

	dirSector := make([]byte, cfbHeaderSize)
	writeDirEntry(dirSector[0:128], "Root Entry", 5, cfbNoStream, cfbNoStream, 1, 3, 512)
	payload := []byte("mini-stream-payload!!")
	writeDirEntry(dirSector[128:256], "Small", 2, cfbNoStream, cfbNoStream, cfbNoStream, 0, uint64(len(payload)))

	miniFatSector := make([]byte, cfbHeaderSize)
	binary.LittleEndian.PutUint32(miniFatSector[0:4], cfbEndOfChain) // mini sector 0
	for i := 1; i < cfbHeaderSize/4; i++ {
		start := i * 4
		binary.LittleEndian.PutUint32(miniFatSector[start:start+4], cfbFreeSector)
	}

	rootMiniData := make([]byte, cfbHeaderSize)
	copy(rootMiniData, payload)

	out := append(header, fatSector...)
	out = append(out, dirSector...)
	out = append(out, miniFatSector...)
	out = append(out, rootMiniData...)
	return out, payload
}

func writeDirEntry(dst []byte, name string, objType byte, left, right, child, startSector uint32, size uint64) {
	for i := range dst {
		dst[i] = 0
	}
	encoded := utf16LE(name)
	nameLen := len(encoded) + 2
	if nameLen > 64 {
		nameLen = 64
	}
	copy(dst[0:nameLen-2], encoded[:nameLen-2])
	dst[nameLen-2] = 0
	dst[nameLen-1] = 0
	binary.LittleEndian.PutUint16(dst[64:66], uint16(nameLen))
	dst[66] = objType
	dst[67] = 1 // black node
	binary.LittleEndian.PutUint32(dst[68:72], left)
	binary.LittleEndian.PutUint32(dst[72:76], right)
	binary.LittleEndian.PutUint32(dst[76:80], child)
	binary.LittleEndian.PutUint32(dst[116:120], startSector)
	binary.LittleEndian.PutUint64(dst[120:128], size)
}

func utf16LE(s string) []byte {
	r := []rune(s)
	u := make([]uint16, len(r))
	for i := range r {
		u[i] = uint16(r[i])
	}
	out := make([]byte, len(u)*2)
	for i, v := range u {
		binary.LittleEndian.PutUint16(out[i*2:i*2+2], v)
	}
	return out
}
