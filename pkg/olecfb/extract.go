package olecfb

import (
	"crypto/sha256"
	"encoding/hex"
	"io"
	"time"

	"github.com/KairuiKin/go-olespec/pkg/oleds"
)

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

	seen := map[string]struct{}{}
	totalBytes := int64(0)
	for _, id := range f.order {
		n, ok := f.nodes[id]
		if !ok {
			continue
		}
		if n.Type != NodeStream {
			continue
		}
		if len(report.Artifacts) >= maxArtifacts {
			report.Partial = true
			report.Warnings = append(report.Warnings, Warning{
				Code:     ErrLimitExceeded,
				Message:  "artifact count limit exceeded",
				Path:     n.Path,
				Offset:   -1,
				Op:       "extract",
				Severity: SeverityWarning,
			})
			break
		}
		if maxArtifactSize > 0 && n.Size > maxArtifactSize {
			report.Partial = true
			report.Warnings = append(report.Warnings, Warning{
				Code:     ErrLimitExceeded,
				Message:  "artifact size limit exceeded",
				Path:     n.Path,
				Offset:   -1,
				Op:       "extract",
				Severity: SeverityWarning,
			})
			continue
		}

		artifact := Artifact{
			ID:           "",
			Kind:         ArtifactStream,
			Status:       ArtifactOK,
			Path:         n.Path,
			MediaType:    "",
			Size:         n.Size,
			Depth:        f.nodeDepth(n.ID),
			SourceNodeID: n.ID,
		}

		sr, err := f.OpenStream(n.Path)
		if err != nil {
			artifact.Status = ArtifactFailed
			artifact.Error = asOLEError(err)
			report.Artifacts = append(report.Artifacts, artifact)
			report.Warnings = append(report.Warnings, warningFromError(err, SeverityWarning))
			continue
		}

		sum, readBytes, isOLE, head, readErr := hashAndProbeStream(sr)
		_ = sr.Close()
		if readErr != nil {
			artifact.Status = ArtifactFailed
			artifact.Error = asOLEError(readErr)
			report.Artifacts = append(report.Artifacts, artifact)
			report.Warnings = append(report.Warnings, warningFromError(readErr, SeverityWarning))
			continue
		}
		artifact.SHA256 = sum
		artifact.ID = "sha256:" + sum
		artifact.Size = readBytes
		if isOLE {
			artifact.Kind = ArtifactOLEFile
		}
		if opt.DetectOLEDS {
			d := oleds.Detect(n.Path, head)
			switch d.Kind {
			case oleds.KindOle10Native, oleds.KindCompObj, oleds.KindPackage:
				artifact.Kind = ArtifactOleObj
				if d.Kind != oleds.KindUnknown {
					artifact.Note = "oleds:" + string(d.Kind)
				}
			}
		}

		if opt.Deduplicate {
			if _, ok := seen[artifact.SHA256]; ok {
				report.Stats.DedupHits++
				continue
			}
			seen[artifact.SHA256] = struct{}{}
		}

		if maxTotalBytes > 0 && totalBytes+artifact.Size > maxTotalBytes {
			report.Partial = true
			report.Warnings = append(report.Warnings, Warning{
				Code:     ErrQuotaExceeded,
				Message:  "total extracted bytes limit exceeded",
				Path:     artifact.Path,
				Offset:   -1,
				Op:       "extract",
				Severity: SeverityWarning,
			})
			break
		}
		totalBytes += artifact.Size
		report.Artifacts = append(report.Artifacts, artifact)
	}

	report.Stats.Duration = time.Since(start)
	report.Stats.ArtifactsTotal = len(report.Artifacts)
	report.Stats.BytesExported = totalBytes
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
