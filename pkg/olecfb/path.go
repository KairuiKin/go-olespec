package olecfb

import (
	"errors"
	"strings"
	"unicode/utf8"
)

type Path string

func ParsePath(s string) (Path, error) {
	if s == "" {
		return "", newError(ErrInvalidArgument, "path is empty", "path.parse", s, -1, nil)
	}
	if s == "/" {
		return Path(s), nil
	}
	if !strings.HasPrefix(s, "/") {
		return "", newError(ErrInvalidArgument, "path must start with /", "path.parse", s, -1, nil)
	}
	if strings.HasSuffix(s, "/") {
		return "", newError(ErrInvalidArgument, "path must not end with /", "path.parse", s, -1, nil)
	}
	raw := strings.Split(s[1:], "/")
	segments := make([]string, 0, len(raw))
	for _, seg := range raw {
		if seg == "" {
			return "", newError(ErrInvalidArgument, "path has empty segment", "path.parse", s, -1, nil)
		}
		decoded, err := DecodeSegment(seg)
		if err != nil {
			return "", newError(ErrInvalidArgument, "invalid path escape", "path.parse", s, -1, err)
		}
		if utf8.RuneCountInString(decoded) > 31 {
			return "", newError(ErrInvalidArgument, "path segment exceeds 31 chars", "path.parse", s, -1, nil)
		}
		segments = append(segments, EncodeSegment(decoded))
	}
	out := "/" + strings.Join(segments, "/")
	if len(out) > 4096 {
		return "", newError(ErrInvalidArgument, "path exceeds 4096 bytes", "path.parse", s, -1, nil)
	}
	return Path(out), nil
}

func JoinPath(parent Path, name string) (Path, error) {
	pp, err := ParsePath(string(parent))
	if err != nil {
		return "", err
	}
	seg := strings.TrimSpace(name)
	if seg == "" {
		return "", newError(ErrInvalidArgument, "name is empty", "path.join", string(parent), -1, nil)
	}
	if utf8.RuneCountInString(seg) > 31 {
		return "", newError(ErrInvalidArgument, "name exceeds 31 chars", "path.join", string(parent), -1, nil)
	}
	enc := EncodeSegment(seg)
	if pp == "/" {
		return ParsePath("/" + enc)
	}
	return ParsePath(string(pp) + "/" + enc)
}

func ParentPath(p Path) Path {
	ps := string(p)
	if ps == "/" {
		return "/"
	}
	i := strings.LastIndex(ps, "/")
	if i <= 0 {
		return "/"
	}
	return Path(ps[:i])
}

func BaseName(p Path) string {
	ps := string(p)
	if ps == "/" {
		return "/"
	}
	i := strings.LastIndex(ps, "/")
	if i < 0 || i+1 >= len(ps) {
		return ""
	}
	name, _ := DecodeSegment(ps[i+1:])
	return name
}

func EncodeSegment(name string) string {
	replaced := strings.ReplaceAll(name, "~", "~0")
	return strings.ReplaceAll(replaced, "/", "~1")
}

func DecodeSegment(seg string) (string, error) {
	var b strings.Builder
	for i := 0; i < len(seg); i++ {
		ch := seg[i]
		if ch != '~' {
			b.WriteByte(ch)
			continue
		}
		if i+1 >= len(seg) {
			return "", errors.New("trailing ~")
		}
		next := seg[i+1]
		switch next {
		case '0':
			b.WriteByte('~')
		case '1':
			b.WriteByte('/')
		default:
			return "", errors.New("invalid escape")
		}
		i++
	}
	return b.String(), nil
}

func CanonicalPath(s string) (Path, error) {
	return ParsePath(s)
}

func PathKey(p Path) string {
	// Placeholder for full Unicode NFC+casefold key; currently lower-case only.
	return strings.ToLower(string(p))
}
