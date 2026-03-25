package olecfb

import (
	"crypto/sha256"
	"encoding/hex"
	"io"
	"strconv"
	"time"

	"github.com/KairuiKin/go-olespec/pkg/oleds"
)

type extractWalker struct {
	report         *ExtractReport
	opt            ExtractOptions
	openOpt        OpenOptions
	seen           map[string]struct{}
	totalBytes     int64
	maxDepth       int
	maxArtifacts   int
	maxTotalBytes  int64
	maxArtifactSize int64
	nextID         int
	indexByID      map[string]int
	stop           bool
}

func (f *File) Extract(opt ExtractOptions) (*ExtractReport, error) {
	if f == nil {
		return nil, newError(ErrInvalidArgument, "file is nil", "extract", "", -1, nil)
	}
	start := time.Now()
	report := &ExtractReport{}

	maxArtifacts := opt.Limits.MaxArtifacts
	if maxArtifacts <= 0 {
		maxArtifacts = 1 << 20
	}
	maxTotalBytes := opt.Limits.MaxTotalBytes
	if maxTotalBytes < 0 {
		maxTotalBytes = 0
	}
	maxArtifactSize := opt.Limits.MaxArtifactSize
	if maxArtifactSize < 0 {
		maxArtifactSize = 0
	}

	openOpt := f.opt
	openOpt.Mode = opt.Mode
	w := &extractWalker{
		report:          report,
		opt:             opt,
		openOpt:         openOpt,
		seen:            map[string]struct{}{},
		maxDepth:        opt.Limits.MaxDepth,
		maxArtifacts:    maxArtifacts,
		maxTotalBytes:   maxTotalBytes,
		maxArtifactSize: maxArtifactSize,
		indexByID:       map[string]int{},
	}
	w.walkFile(f, "", "", 0)

	report.Stats.Duration = time.Since(start)
	report.Stats.ArtifactsTotal = len(report.Artifacts)
	report.Stats.BytesExported = w.totalBytes
	for _, a := range report.Artifacts {
		switch a.Status {
		case ArtifactOK:
			report.Stats.ArtifactsOK++
		case ArtifactPartial:
			report.Stats.ArtifactsPartial++
		case ArtifactFailed:
			report.Stats.ArtifactsFailed++
		}
		if a.Depth > report.Stats.MaxDepthReached {
			report.Stats.MaxDepthReached = a.Depth
		}
	}
	report.Warnings = append(report.Warnings, f.report.Warnings...)
	report.Repairs = append(report.Repairs, f.report.Repairs...)
	if f.report.Partial {
		report.Partial = true
	}
	return report, nil
}

func (w *extractWalker) walkFile(file *File, pathPrefix, parentID string, depth int) {
	if w.stop || file == nil {
		return
	}
	for _, id := range file.order {
		if w.stop {
			return
		}
		n, ok := file.nodes[id]
		if !ok || n.Type != NodeStream {
			continue
		}
		w.walkStream(file, n, pathPrefix, parentID, depth)
	}
}

func (w *extractWalker) walkStream(file *File, n Node, pathPrefix, parentID string, depth int) {
	if len(w.report.Artifacts) >= w.maxArtifacts {
		w.report.Partial = true
		w.report.Warnings = append(w.report.Warnings, Warning{
			Code:     ErrLimitExceeded,
			Message:  "artifact count limit exceeded",
			Path:     n.Path,
			Offset:   -1,
			Op:       "extract",
			Severity: SeverityWarning,
		})
		w.stop = true
		return
	}
	if w.maxArtifactSize > 0 && n.Size > w.maxArtifactSize {
		w.report.Partial = true
		w.report.Warnings = append(w.report.Warnings, Warning{
			Code:     ErrLimitExceeded,
			Message:  "artifact size limit exceeded",
			Path:     n.Path,
			Offset:   -1,
			Op:       "extract",
			Severity: SeverityWarning,
		})
		return
	}

	artifactPath := joinArtifactPath(pathPrefix, n.Path)
	artifact := Artifact{
		Kind:         ArtifactStream,
		Status:       ArtifactOK,
		Path:         artifactPath,
		Size:         n.Size,
		Depth:        depth,
		ParentID:     parentID,
		SourceNodeID: n.ID,
	}

	sr, err := file.OpenStream(n.Path)
	if err != nil {
		artifact.Status = ArtifactFailed
		artifact.Error = asOLEError(err)
		w.appendArtifact(artifact)
		w.report.Warnings = append(w.report.Warnings, warningFromError(err, SeverityWarning))
		return
	}

	sum, readBytes, isOLE, head, readErr := hashAndProbeStream(sr)
	_ = sr.Close()
	if readErr != nil {
		artifact.Status = ArtifactFailed
		artifact.Error = asOLEError(readErr)
		w.appendArtifact(artifact)
		w.report.Warnings = append(w.report.Warnings, warningFromError(readErr, SeverityWarning))
		return
	}
	artifact.SHA256 = sum
	artifact.Size = readBytes
	artifact.ID = w.newArtifactID(sum)
	if isOLE {
		artifact.Kind = ArtifactOLEFile
	}
	if w.opt.DetectImages {
		if media, ok := detectImageMedia(head); ok {
			artifact.Kind = ArtifactImage
			artifact.MediaType = media
		}
	}
	if w.opt.DetectOLEDS {
		d := oleds.Detect(n.Path, head)
		switch d.Kind {
		case oleds.KindOle10Native, oleds.KindCompObj, oleds.KindPackage:
			artifact.Kind = ArtifactOleObj
			artifact.Note = "oleds:" + string(d.Kind)
		}
	}

	if w.opt.Deduplicate {
		if _, ok := w.seen[artifact.SHA256]; ok {
			w.report.Stats.DedupHits++
			return
		}
		w.seen[artifact.SHA256] = struct{}{}
	}
	if w.maxTotalBytes > 0 && w.totalBytes+artifact.Size > w.maxTotalBytes {
		w.report.Partial = true
		w.report.Warnings = append(w.report.Warnings, Warning{
			Code:     ErrQuotaExceeded,
			Message:  "total extracted bytes limit exceeded",
			Path:     artifact.Path,
			Offset:   -1,
			Op:       "extract",
			Severity: SeverityWarning,
		})
		w.stop = true
		return
	}
	needsPayload := w.opt.IncludeRaw || isOLE
	var payload []byte
	if needsPayload {
		payload, err = readStreamAll(file, n.Path, w.maxArtifactSize)
		if err != nil {
			w.report.Partial = true
			w.report.Warnings = append(w.report.Warnings, warningFromError(err, SeverityWarning))
			return
		}
		if w.opt.IncludeRaw {
			artifact.Raw = append([]byte(nil), payload...)
		}
	}

	w.totalBytes += artifact.Size
	w.appendArtifact(artifact)
	if !isOLE || w.stop {
		return
	}

	if w.maxDepth > 0 && depth >= w.maxDepth {
		w.report.Partial = true
		w.report.Warnings = append(w.report.Warnings, Warning{
			Code:     ErrDepthExceeded,
			Message:  "extract depth limit reached",
			Path:     artifact.Path,
			Offset:   -1,
			Op:       "extract",
			Severity: SeverityWarning,
		})
		return
	}

	if payload == nil {
		payload, err = readStreamAll(file, n.Path, w.maxArtifactSize)
		if err != nil {
			w.report.Partial = true
			w.report.Warnings = append(w.report.Warnings, warningFromError(err, SeverityWarning))
			return
		}
	}
	nested, err := OpenBytes(payload, w.openOpt)
	if err != nil {
		w.report.Partial = true
		w.report.Warnings = append(w.report.Warnings, warningFromError(err, SeverityWarning))
		return
	}
	defer nested.Close()
	w.walkFile(nested, artifact.Path, artifact.ID, depth+1)
}

func (w *extractWalker) appendArtifact(a Artifact) {
	w.report.Artifacts = append(w.report.Artifacts, a)
	idx := len(w.report.Artifacts) - 1
	w.indexByID[a.ID] = idx
	if a.ParentID == "" {
		return
	}
	pIdx, ok := w.indexByID[a.ParentID]
	if !ok {
		return
	}
	parent := w.report.Artifacts[pIdx]
	parent.Children++
	w.report.Artifacts[pIdx] = parent
}

func (w *extractWalker) newArtifactID(sha string) string {
	w.nextID++
	return "sha256:" + sha + "#" + strconv.Itoa(w.nextID)
}

func joinArtifactPath(prefix, path string) string {
	if prefix == "" {
		return path
	}
	return prefix + "!" + path
}

func readStreamAll(f *File, streamPath string, maxBytes int64) ([]byte, error) {
	sr, err := f.OpenStream(streamPath)
	if err != nil {
		return nil, err
	}
	defer sr.Close()
	if maxBytes > 0 && sr.Size() > maxBytes {
		return nil, newError(ErrLimitExceeded, "artifact size limit exceeded", "extract.read_stream", streamPath, -1, nil)
	}
	data, err := io.ReadAll(sr)
	if err != nil {
		return nil, err
	}
	return data, nil
}

func hashAndProbeStream(r io.Reader) (sha string, total int64, isOLE bool, head []byte, err error) {
	h := sha256.New()
	buf := make([]byte, 32*1024)
	var first8 [8]byte
	firstRead := 0
	head = make([]byte, 0, 64*1024)
	for {
		n, e := r.Read(buf)
		if n > 0 {
			if firstRead < 8 {
				copyN := n
				if copyN > 8-firstRead {
					copyN = 8 - firstRead
				}
				copy(first8[firstRead:firstRead+copyN], buf[:copyN])
				firstRead += copyN
			}
			if _, wErr := h.Write(buf[:n]); wErr != nil {
				return "", total, false, nil, wErr
			}
			total += int64(n)
			if len(head) < cap(head) {
				remain := cap(head) - len(head)
				copyN := n
				if copyN > remain {
					copyN = remain
				}
				head = append(head, buf[:copyN]...)
			}
		}
		if e == io.EOF {
			break
		}
		if e != nil {
			return "", total, false, nil, e
		}
	}
	isOLE = firstRead == 8 && first8 == cfbSignature
	return hex.EncodeToString(h.Sum(nil)), total, isOLE, head, nil
}

func asOLEError(err error) *OLEError {
	if err == nil {
		return nil
	}
	if oe, ok := err.(*OLEError); ok {
		return oe
	}
	return &OLEError{
		Code:    ErrUnsupported,
		Message: err.Error(),
		Offset:  -1,
	}
}

func detectImageMedia(head []byte) (string, bool) {
	if len(head) >= 8 &&
		head[0] == 0x89 && head[1] == 0x50 && head[2] == 0x4E && head[3] == 0x47 &&
		head[4] == 0x0D && head[5] == 0x0A && head[6] == 0x1A && head[7] == 0x0A {
		return "image/png", true
	}
	if len(head) >= 3 && head[0] == 0xFF && head[1] == 0xD8 && head[2] == 0xFF {
		return "image/jpeg", true
	}
	if len(head) >= 6 {
		if string(head[:6]) == "GIF87a" || string(head[:6]) == "GIF89a" {
			return "image/gif", true
		}
	}
	if len(head) >= 2 && head[0] == 0x42 && head[1] == 0x4D {
		return "image/bmp", true
	}
	if len(head) >= 4 {
		if (head[0] == 0x49 && head[1] == 0x49 && head[2] == 0x2A && head[3] == 0x00) ||
			(head[0] == 0x4D && head[1] == 0x4D && head[2] == 0x00 && head[3] == 0x2A) {
			return "image/tiff", true
		}
	}
	if len(head) >= 12 && string(head[:4]) == "RIFF" && string(head[8:12]) == "WEBP" {
		return "image/webp", true
	}
	return "", false
}
