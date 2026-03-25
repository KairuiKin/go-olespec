package olecfb

import (
	"testing"

	"github.com/KairuiKin/go-olespec/pkg/oleps"
)

func TestTxPutPropertySet(t *testing.T) {
	ps := buildSummaryPropertySetStreamBytes("Core Title", 9)
	fileBytes, _ := buildValidV4FileWithNamedStream("\x05SummaryInformation", ps)
	f, err := OpenBytesRW(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytesRW returned error: %v", err)
	}

	stream, err := f.OpenPropertySet("/\x05SummaryInformation")
	if err != nil {
		t.Fatalf("OpenPropertySet returned error: %v", err)
	}
	set, ok := stream.SummaryInformation()
	if !ok {
		t.Fatal("SummaryInformation set not found")
	}
	set.SetString(oleps.PIDTitle, "Edited Title")

	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.PutPropertySet("/\x05SummaryInformation", stream); err != nil {
		t.Fatalf("PutPropertySet returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}

	set2, err := f.OpenSummaryInformation()
	if err != nil {
		t.Fatalf("OpenSummaryInformation returned error: %v", err)
	}
	title, ok := set2.GetString(oleps.PIDTitle)
	if !ok {
		t.Fatal("title property not found")
	}
	if title != "Edited Title" {
		t.Fatalf("unexpected title after writeback: %q", title)
	}
}

func TestTxPutSummaryInformation(t *testing.T) {
	ps := buildSummaryPropertySetStreamBytes("Core Title", 9)
	fileBytes, _ := buildValidV4FileWithNamedStream("\x05SummaryInformation", ps)
	f, err := OpenBytesRW(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytesRW returned error: %v", err)
	}
	set, err := f.OpenSummaryInformation()
	if err != nil {
		t.Fatalf("OpenSummaryInformation returned error: %v", err)
	}
	set.SetString(oleps.PIDTitle, "Edited via convenience")
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.PutSummaryInformation(set); err != nil {
		t.Fatalf("PutSummaryInformation returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	set2, err := f.OpenSummaryInformation()
	if err != nil {
		t.Fatalf("OpenSummaryInformation returned error: %v", err)
	}
	title, ok := set2.GetString(oleps.PIDTitle)
	if !ok || title != "Edited via convenience" {
		t.Fatalf("unexpected title after convenience writeback: %q", title)
	}
}

func TestTxPutDocumentSummaryInformation(t *testing.T) {
	ps := buildDocumentSummaryPropertySetStreamBytes("Doc Author")
	fileBytes, _ := buildValidV4FileWithNamedStream("\x05DocumentSummaryInformation", ps)
	f, err := OpenBytesRW(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytesRW returned error: %v", err)
	}
	set, err := f.OpenDocumentSummaryInformation()
	if err != nil {
		t.Fatalf("OpenDocumentSummaryInformation returned error: %v", err)
	}
	set.SetString(oleps.PIDAuthor, "Edited Author")
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.PutDocumentSummaryInformation(set); err != nil {
		t.Fatalf("PutDocumentSummaryInformation returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	set2, err := f.OpenDocumentSummaryInformation()
	if err != nil {
		t.Fatalf("OpenDocumentSummaryInformation returned error: %v", err)
	}
	author, ok := set2.GetString(oleps.PIDAuthor)
	if !ok || author != "Edited Author" {
		t.Fatalf("unexpected author after convenience writeback: %q", author)
	}
}

func TestTxPutSummaryInformation_ExistingInvalidStreamFails(t *testing.T) {
	fileBytes, _ := buildValidV4FileWithNamedStream("\x05SummaryInformation", []byte("not-a-property-set"))
	f, err := OpenBytesRW(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytesRW returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	set := &oleps.PropertySet{
		Properties: map[uint32]oleps.Property{
			oleps.PIDTitle: {ID: oleps.PIDTitle, Type: oleps.VTLPWSTR, Value: "X"},
		},
		Order: []uint32{oleps.PIDTitle},
	}
	err = tx.PutSummaryInformation(set)
	if err == nil {
		t.Fatal("expected parse error for existing invalid property stream")
	}
	if !IsCode(err, ErrDirCorrupt) {
		t.Fatalf("expected ErrDirCorrupt, got %v", err)
	}
}

func TestTxPutSummaryInformation_PreserveOtherSets(t *testing.T) {
	stream := &oleps.Stream{
		ByteOrder: 0xFFFE,
		Sets: []oleps.PropertySet{
			{
				FormatID: oleps.FMTIDSummaryInformation,
				Properties: map[uint32]oleps.Property{
					oleps.PIDTitle: {
						ID:    oleps.PIDTitle,
						Type:  oleps.VTLPWSTR,
						Value: "Old Title",
					},
				},
				Order: []uint32{oleps.PIDTitle},
			},
			{
				FormatID: oleps.FMTIDDocumentSummaryInformation,
				Properties: map[uint32]oleps.Property{
					oleps.PIDAuthor: {
						ID:    oleps.PIDAuthor,
						Type:  oleps.VTLPWSTR,
						Value: "Keep Author",
					},
				},
				Order: []uint32{oleps.PIDAuthor},
			},
		},
	}
	data, err := oleps.Marshal(stream)
	if err != nil {
		t.Fatalf("Marshal returned error: %v", err)
	}
	fileBytes, _ := buildValidV4FileWithNamedStream("\x05SummaryInformation", data)
	f, err := OpenBytesRW(fileBytes, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytesRW returned error: %v", err)
	}

	set, err := f.OpenSummaryInformation()
	if err != nil {
		t.Fatalf("OpenSummaryInformation returned error: %v", err)
	}
	set.SetString(oleps.PIDTitle, "New Title")

	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.PutSummaryInformation(set); err != nil {
		t.Fatalf("PutSummaryInformation returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}

	updated, err := f.OpenPropertySet("/\x05SummaryInformation")
	if err != nil {
		t.Fatalf("OpenPropertySet returned error: %v", err)
	}
	if len(updated.Sets) != 2 {
		t.Fatalf("expected 2 sets, got %d", len(updated.Sets))
	}
	si, ok := updated.SummaryInformation()
	if !ok {
		t.Fatal("summary set missing")
	}
	title, ok := si.GetString(oleps.PIDTitle)
	if !ok || title != "New Title" {
		t.Fatalf("unexpected title: %q", title)
	}
	dsi, ok := updated.DocumentSummaryInformation()
	if !ok {
		t.Fatal("document summary set missing")
	}
	author, ok := dsi.GetString(oleps.PIDAuthor)
	if !ok || author != "Keep Author" {
		t.Fatalf("unexpected author: %q", author)
	}
}
