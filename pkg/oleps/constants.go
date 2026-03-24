package oleps

var (
	// F29F85E0-4FF9-1068-AB91-08002B27B3D9
	FMTIDSummaryInformation = GUID{0xE0, 0x85, 0x9F, 0xF2, 0xF9, 0x4F, 0x68, 0x10, 0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9}
	// D5CDD502-2E9C-101B-9397-08002B2CF9AE
	FMTIDDocumentSummaryInformation = GUID{0x02, 0xD5, 0xCD, 0xD5, 0x9C, 0x2E, 0x1B, 0x10, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE}
)

const (
	PIDCodePage     uint32 = 1
	PIDTitle        uint32 = 2
	PIDSubject      uint32 = 3
	PIDAuthor       uint32 = 4
	PIDKeywords     uint32 = 5
	PIDComments     uint32 = 6
	PIDTemplate     uint32 = 7
	PIDLastAuthor   uint32 = 8
	PIDRevNumber    uint32 = 9
	PIDEditTime     uint32 = 10
	PIDLastPrinted  uint32 = 11
	PIDCreateTime   uint32 = 12
	PIDLastSaveTime uint32 = 13
	PIDPageCount    uint32 = 14
	PIDWordCount    uint32 = 15
	PIDCharCount    uint32 = 16
	PIDAppName      uint32 = 18
	PIDSecurity     uint32 = 19
)
