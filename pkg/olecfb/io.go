package olecfb

import "io"

func readFullAt(readAt func([]byte, int64) (int, error), buf []byte, off int64) error {
	read := 0
	for read < len(buf) {
		n, err := readAt(buf[read:], off+int64(read))
		if n > 0 {
			read += n
		}
		if err != nil {
			if err == io.EOF && read == len(buf) {
				return nil
			}
			return err
		}
		if n == 0 {
			return io.ErrUnexpectedEOF
		}
	}
	return nil
}
