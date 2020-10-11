# -*- coding: utf-8 -*-

# https://github.com/decalage2/oletools
# https://github.com/decalage2/olefile
# https://github.com/rembish/cfb
# https://github.com/decalage2/oletools/wiki/olemeta

import olefile
import struct
import datetime
import io

VT_EMPTY=0; VT_NULL=1; VT_I2=2; VT_I4=3; VT_R4=4; VT_R8=5; VT_CY=6;
VT_DATE=7; VT_BSTR=8; VT_DISPATCH=9; VT_ERROR=10; VT_BOOL=11;
VT_VARIANT=12; VT_UNKNOWN=13; VT_DECIMAL=14; VT_I1=16; VT_UI1=17;
VT_UI2=18; VT_UI4=19; VT_I8=20; VT_UI8=21; VT_INT=22; VT_UINT=23;
VT_VOID=24; VT_HRESULT=25; VT_PTR=26; VT_SAFEARRAY=27; VT_CARRAY=28;
VT_USERDEFINED=29; VT_LPSTR=30; VT_LPWSTR=31; VT_FILETIME=64;
VT_BLOB=65; VT_STREAM=66; VT_STORAGE=67; VT_STREAMED_OBJECT=68;
VT_STORED_OBJECT=69; VT_BLOB_OBJECT=70; VT_CF=71; VT_CLSID=72;
VT_VECTOR=0x1000;

SUMMARY_ATTRIBS = ['codepage', 'title', 'subject', 'author', 'keywords', 'comments',
        'template', 'last_saved_by', 'revision_number', 'total_edit_time',
        'last_printed', 'create_time', 'last_saved_time', 'num_pages',
        'num_words', 'num_chars', 'thumbnail', 'creating_application',
        'security']

DOCSUM_ATTRIBS = ['codepage_doc', 'category', 'presentation_target', 'bytes', 'lines', 'paragraphs',
        'slides', 'notes', 'hidden_slides', 'mm_clips',
        'scale_crop', 'heading_pairs', 'titles_of_parts', 'manager',
        'company', 'links_dirty', 'chars_with_spaces', 'unused', 'shared_doc',
        'link_base', 'hlinks', 'hlinks_changed', 'version', 'dig_sig',
        'content_type', 'content_status', 'language', 'doc_version']

def u32(x): return struct.unpack("<L", x)[0]

# https://github.com/decalage2/olefile/blob/d344bc25e0f6f549e1b6d48cf40af76cdac70f9b/olefile/olefile.py#L320
def i8(c):
    return c if c.__class__ is int else c[0]

def i16(c, o = 0):
    """
    Converts a 2-bytes (16 bits) string to an integer.
    :param c: string containing bytes to convert
    :param o: offset of bytes to convert in string
    """
    return struct.unpack("<H", c[o:o+2])[0]

def i32(c, o = 0):
    """
    Converts a 4-bytes (32 bits) string to an integer.
    :param c: string containing bytes to convert
    :param o: offset of bytes to convert in string
    """
    return struct.unpack("<I", c[o:o+4])[0]

def _clsid(clsid):
    #print((i32(clsid, 0), i16(clsid, 4), i16(clsid, 6)) +
    #        tuple(map(i8, clsid[8:16])))
    """
    Converts a CLSID to a human-readable string.
    :param clsid: string of length 16.
    """
    assert len(clsid) == 16
    # if clsid is only made of null bytes, return an empty string:
    # (PL: why not simply return the string with zeroes?)
    if not clsid.strip(b"\0"):
        return ""
    return ((i32(clsid, 0), i16(clsid, 4), i16(clsid, 6)) +
            tuple(map(i8, clsid[8:16])))

    #return (("%08X-%04X-%04X-%02X%02X-" + "%02X" * 6) %
    #        ((i32(clsid, 0), i16(clsid, 4), i16(clsid, 6)) +
    #        tuple(map(i8, clsid[8:16]))))

class CFBF:
    def __init__(self):
        self.SummaryInfo_data_idx = -1
        self.DocumentSummaryInfo_data_idx = -1

    def getAttrDecodeString(self, name):
        try:
            return getattr(self.meta, name).decode("cp949")
        except:
            pass

        try:
            return getattr(self.meta, name).decode("utf-8")
        except:
            pass
        
        try:
            return getattr(self.meta, name)
        except:
            return ""

    def decodeString(self, name):
        try:
            return self.meta[name].decode("cp949")
        except:
            pass

        try:
            return self.meta[name].decode("utf-8")
        except:
            pass
        
        try:
            return self.meta[name]
        except:
            return ""

    # https://github.com/decalage2/oletools/blob/master/oletools/olemeta.py
    def ole_meta(self, ole):
        self.result_meta = {}
        self.meta = ole.get_metadata()

        for i in SUMMARY_ATTRIBS + DOCSUM_ATTRIBS:
            data = self.getAttrDecodeString(i)
            if data != "" and data != None and i != "thumbnail":
                self.result_meta[i] = data

        return self.result_meta

    # https://github.com/decalage2/olefile/blob/d344bc25e0f6f549e1b6d48cf40af76cdac70f9b/olefile/olefile.py#L1153
    def _decode_utf16_str(self, utf16_str, errors='replace'):
        unicode_str = utf16_str.decode('UTF-16LE', errors)
        if self.path_encoding:
            # an encoding has been specified for path names:
            return unicode_str.encode(self.path_encoding, errors)
        else:
            # path_encoding=None, return the Unicode string as-is:
            return unicode_str

    def getproperties(self, fp, desc, no_conversion=None):
        # https://github.com/decalage2/olefile/blob/d344bc25e0f6f549e1b6d48cf40af76cdac70f9b/olefile/olefile.py#L511
        # https://github.com/decalage2/olefile/blob/d344bc25e0f6f549e1b6d48cf40af76cdac70f9b/olefile/olefile.py#L520
        convert_time = True

        if no_conversion == None:
            no_conversion = []

        data = {}
        try:
            # header
            s = fp.read(28)
            clsid = _clsid(s[8:24])
            # format id
            s = fp.read(20)
            fmtid = _clsid(s[:16])
            fp.seek(i32(s, 16))
            # get section
            s = b"****" + fp.read(i32(fp.read(4))-4)
            # number of properties:
            num_props = i32(s, 4)
        except BaseException as exc:
            #print("{} : {}".format(desc, self.path))
            return data

        num_props = min(num_props, int(len(s) / 8))
        for i in xrange(num_props):
            property_id = 0

            try:
                property_id = i32(s, 8+i*8)
                offset = i32(s, 12+i*8)
                property_type = i32(s, offset)

                if property_type == VT_I2: # 16-bit signed integer
                    value = i16(s, offset+4)
                    if value >= 32768:
                        value = value - 65536
                elif property_type == VT_UI2: # 2-byte unsigned integer
                    value = i16(s, offset+4)
                elif property_type in (VT_I4, VT_INT, VT_ERROR):
                        # VT_I4: 32-bit signed integer
                        # VT_ERROR: HRESULT, similar to 32-bit signed integer,
                        # see https://msdn.microsoft.com/en-us/library/cc230330.aspx
                    value = i32(s, offset+4)
                elif property_type in (VT_UI4, VT_UINT): # 4-byte unsigned integer
                    value = i32(s, offset+4) # FIXME
                elif property_type in (VT_BSTR, VT_LPSTR):
                        # CodePageString, see https://msdn.microsoft.com/en-us/library/dd942354.aspx
                        # size is a 32 bits integer, including the null terminator, and
                        # possibly trailing or embedded null chars
                        # TODO: if codepage is unicode, the string should be converted as such
                    count = i32(s, offset+4)
                    value = s[offset+8:offset+8+count-1]
                        # remove all null chars:
                    value = value.replace(b'\x00', b'')
                elif property_type == VT_BLOB:
                        # binary large object (BLOB)
                        # see https://msdn.microsoft.com/en-us/library/dd942282.aspx
                    count = i32(s, offset+4)
                    value = s[offset+8:offset+8+count]
                elif property_type == VT_LPWSTR:
                        # UnicodeString
                        # see https://msdn.microsoft.com/en-us/library/dd942313.aspx
                        # "the string should NOT contain embedded or additional trailing
                        # null characters."
                    count = i32(s, offset+4)
                    value = self._decode_utf16_str(s[offset+8:offset+8+count*2])
                elif property_type == VT_FILETIME:
                    value = int(i32(s, offset+4)) + (int(i32(s, offset+8))<<32)
                        # FILETIME is a 64-bit int: "number of 100ns periods
                        # since Jan 1,1601".
                    if convert_time and property_id not in no_conversion:
                            ##log.debug('Converting property #%d to python datetime, value=%d=%fs'
                            ##        %(property_id, value, float(value)/10000000))
                            # convert FILETIME to Python datetime.datetime
                            # inspired from https://code.activestate.com/recipes/511425-filetime-to-datetime/
                        _FILETIME_null_date = datetime.datetime(1601, 1, 1, 0, 0, 0)
                            ##log.debug('timedelta days=%d' % (value//(10*1000000*3600*24)))
                        value = _FILETIME_null_date + datetime.timedelta(microseconds=value//10)
                    else:
                            # legacy code kept for backward compatibility: returns a
                            # number of seconds since Jan 1,1601
                        value = value // 10000000 # seconds
                elif property_type == VT_UI1: # 1-byte unsigned integer
                    value = i8(s[offset+4])
                elif property_type == VT_CLSID:
                    value = _clsid(s[offset+4:offset+20])
                elif property_type == VT_CF:
                        # PropertyIdentifier or ClipboardData??
                        # see https://msdn.microsoft.com/en-us/library/dd941945.aspx
                    count = i32(s, offset+4)
                    value = s[offset+8:offset+8+count]
                elif property_type == VT_BOOL:
                        # VARIANT_BOOL, 16 bits bool, 0x0000=Fals, 0xFFFF=True
                        # see https://msdn.microsoft.com/en-us/library/cc237864.aspx
                    value = bool(i16(s, offset+4))
                else:
                    value = None # everything else yields "None"
                        ##log.debug('property id=%d: type=%d not implemented in parser yet' % (property_id, property_type))

                    # missing: VT_EMPTY, VT_NULL, VT_R4, VT_R8, VT_CY, VT_DATE,
                    # VT_DECIMAL, VT_I1, VT_I8, VT_UI8,
                    # see https://msdn.microsoft.com/en-us/library/dd942033.aspx

                    # FIXME: add support for VT_VECTOR
                    # VT_VECTOR is a 32 uint giving the number of items, followed by
                    # the items in sequence. The VT_VECTOR value is combined with the
                    # type of items, e.g. VT_VECTOR|VT_BSTR
                    # see https://msdn.microsoft.com/en-us/library/dd942011.aspx

                    # print("%08x" % property_id, repr(value), end=" ")
                    # print("(%s)" % VT[i32(s, offset) & 0xFFF])

                data[property_id] = value
            except BaseException as exc:
                pass

        return data

    def parseData(self, data, idx, desc, no_conversion=None):
        # https://github.com/decalage2/olefile/blob/ced013f9c20c59ffebdd8f963f09b6f229d79b7a/olefile/olefile.py#L763
        StartSectorID = u32(data[idx+0x74:idx+0x78]) # self.isectStart
        Size = u32(data[idx+0x78:idx+0x7c]) # self.sizeLow

        sectorIdx = (StartSectorID + 1) * 512
        sectorData = data[sectorIdx:sectorIdx+Size]

        # https://github.com/decalage2/olefile/blob/d344bc25e0f6f549e1b6d48cf40af76cdac70f9b/olefile/olefile.py#L2130
        return self.getproperties(io.BytesIO(sectorData), desc, no_conversion)

    def parseMetadata(self, path):
        self.meta = {}
        self.result_meta = {}

        try:
            with open(path, "rb") as f:
                data = f.read()
                SumInfoAddr = data.find("\x05SummaryInformation".encode("utf-16-le"))
                DocSumInfoAddr = data.find("\x05DocumentSummaryInformation".encode("utf-16-le"))

                # https://github.com/decalage2/olefile/blob/d344bc25e0f6f549e1b6d48cf40af76cdac70f9b/olefile/olefile.py#L494
                if SumInfoAddr != -1:
                    props = self.parseData(data, SumInfoAddr, desc = "\x05SummaryInformation", no_conversion=[10])
                    for i in range(len(SUMMARY_ATTRIBS)):
                        value = props.get(i+1, None)
                        self.meta[SUMMARY_ATTRIBS[i]] = value

                if DocSumInfoAddr != -1:
                    props = self.parseData(data, DocSumInfoAddr, desc = "\x05DocumentSummaryInformation")
                    for i in range(len(DOCSUM_ATTRIBS)):
                        value = props.get(i+1, None)
                        self.meta[DOCSUM_ATTRIBS[i]] = value
        
        except Exception as e:
            #print("[*] parseMetadata Error {}".format(self.path))
            #print(e)
            pass
        
        finally:
            self.SummaryInfo_data_idx = -1
            self.DocumentSummaryInfo_data_idx = -1

        for i in SUMMARY_ATTRIBS + DOCSUM_ATTRIBS:
            data = self.decodeString(i)
            if data != "" and data != None and i != "thumbnail":
                self.result_meta[i] = data
                
        return self.result_meta

    def run(self, path):
        self.path = path
        try:
            ole = olefile.OleFileIO(self.path)
            return self.ole_meta(ole)
        except Exception:
            #print("[*] olefile Error {}".format(path))
            return self.parseMetadata(self.path)