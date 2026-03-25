package olecfb

import (
	"bytes"
	"reflect"
	"testing"
)

func TestWalkDeterministic(t *testing.T) {
	f := buildDeterministicTreeFile(t)

	paths1 := collectWalkPaths(t, f)
	paths2 := collectWalkPaths(t, f)
	if !reflect.DeepEqual(paths1, paths2) {
		t.Fatalf("walk output is not deterministic: %#v vs %#v", paths1, paths2)
	}
}

func TestWalkExDeterministic(t *testing.T) {
	f := buildDeterministicTreeFile(t)

	dfs1 := collectWalkExPaths(t, f, WalkDFS)
	dfs2 := collectWalkExPaths(t, f, WalkDFS)
	if !reflect.DeepEqual(dfs1, dfs2) {
		t.Fatalf("walkex dfs output is not deterministic: %#v vs %#v", dfs1, dfs2)
	}

	bfs1 := collectWalkExPaths(t, f, WalkBFS)
	bfs2 := collectWalkExPaths(t, f, WalkBFS)
	if !reflect.DeepEqual(bfs1, bfs2) {
		t.Fatalf("walkex bfs output is not deterministic: %#v vs %#v", bfs1, bfs2)
	}
}

func buildDeterministicTreeFile(t *testing.T) *File {
	t.Helper()
	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	// Intentionally unsorted insertion order.
	if err := tx.CreateStorage("/zeta"); err != nil {
		t.Fatalf("CreateStorage returned error: %v", err)
	}
	if err := tx.CreateStorage("/Alpha"); err != nil {
		t.Fatalf("CreateStorage returned error: %v", err)
	}
	if err := tx.CreateStorage("/beta"); err != nil {
		t.Fatalf("CreateStorage returned error: %v", err)
	}
	if err := tx.PutStream("/Alpha/s1", bytes.NewReader([]byte("1")), 1); err != nil {
		t.Fatalf("PutStream returned error: %v", err)
	}
	if err := tx.PutStream("/zeta/s2", bytes.NewReader([]byte("2")), 1); err != nil {
		t.Fatalf("PutStream returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	return f
}

func collectWalkPaths(t *testing.T, f *File) []string {
	t.Helper()
	var out []string
	if err := f.Walk(func(n Node) error {
		out = append(out, n.Path)
		return nil
	}); err != nil {
		t.Fatalf("Walk returned error: %v", err)
	}
	return out
}

func collectWalkExPaths(t *testing.T, f *File, order WalkOrder) []string {
	t.Helper()
	var out []string
	_, err := f.WalkEx(WalkOptions{IncludeRoot: true, Order: order}, func(ev WalkEvent) error {
		out = append(out, ev.Node.Path)
		return nil
	})
	if err != nil {
		t.Fatalf("WalkEx returned error: %v", err)
	}
	return out
}
