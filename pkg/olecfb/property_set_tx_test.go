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
	set.Properties[oleps.PIDTitle] = oleps.Property{
		ID:    oleps.PIDTitle,
		Type:  oleps.VTLPWSTR,
		Value: "Edited Title",
	}

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
	set.Properties[oleps.PIDTitle] = oleps.Property{
		ID:    oleps.PIDTitle,
		Type:  oleps.VTLPWSTR,
		Value: "Edited via convenience",
	}
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
	set.Properties[oleps.PIDAuthor] = oleps.Property{
		ID:    oleps.PIDAuthor,
		Type:  oleps.VTLPWSTR,
		Value: "Edited Author",
	}
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
