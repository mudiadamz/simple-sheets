"""Open password-protected Excel files (.xlsx / .xlsm saved with "Encrypt with password") using only the
standard library: an OLE compound-file reader, ECMA-376 Agile and Standard encryption, and AES.
decrypt(path, password) -> the plain .xlsx bytes."""
import base64
import hashlib
import struct
import xml.etree.ElementTree as ET

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


class WrongPassword(Exception):
    pass


class NotEncrypted(Exception):
    pass


# ---------- AES (decryption, table driven) ----------

def _xt(a):
    return ((a << 1) ^ 0x1B) & 0xFF if a & 0x80 else a << 1


def _mul(a, b):
    r = 0
    while b:
        if b & 1:
            r ^= a
        a, b = _xt(a), b >> 1
    return r


SBOX = [0] * 256
INV = [0] * 256
_p = _q = 1
while True:   # standard S-box construction
    _p = _p ^ ((_p << 1) & 0xFF) ^ (0x1B if _p & 0x80 else 0)
    _q ^= _q << 1
    _q ^= _q << 2
    _q ^= _q << 4
    _q &= 0xFF
    if _q & 0x80:
        _q ^= 0x09
    _x = _q ^ ((_q << 1) | (_q >> 7)) & 0xFF ^ ((_q << 2) | (_q >> 6)) & 0xFF ^ ((_q << 3) | (_q >> 5)) & 0xFF \
        ^ ((_q << 4) | (_q >> 4)) & 0xFF ^ 0x63
    SBOX[_p] = _x & 0xFF
    if _p == 1:
        break
SBOX[0] = 0x63
for _i, _v in enumerate(SBOX):
    INV[_v] = _i
D0 = [(_mul(INV[i], 14) << 24) | (_mul(INV[i], 9) << 16) | (_mul(INV[i], 13) << 8) | _mul(INV[i], 11) for i in range(256)]
D1 = [((x >> 8) | (x << 24)) & 0xFFFFFFFF for x in D0]
D2 = [((x >> 16) | (x << 16)) & 0xFFFFFFFF for x in D0]
D3 = [((x >> 24) | (x << 8)) & 0xFFFFFFFF for x in D0]
RCON = [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80, 0x1B, 0x36]
E0 = [(_mul(SBOX[i], 2) << 24) | (SBOX[i] << 16) | (SBOX[i] << 8) | _mul(SBOX[i], 3) for i in range(256)]
E1 = [((x >> 8) | (x << 24)) & 0xFFFFFFFF for x in E0]
E2 = [((x >> 16) | (x << 16)) & 0xFFFFFFFF for x in E0]
E3 = [((x >> 24) | (x << 8)) & 0xFFFFFFFF for x in E0]


def _expand(key):
    nk = len(key) // 4
    rounds = nk + 6
    w = [int.from_bytes(key[4 * i:4 * i + 4], "big") for i in range(nk)]
    for i in range(nk, 4 * (rounds + 1)):
        t = w[i - 1]
        if i % nk == 0:
            t = ((t << 8) | (t >> 24)) & 0xFFFFFFFF
            t = (SBOX[t >> 24] << 24) | (SBOX[(t >> 16) & 255] << 16) | (SBOX[(t >> 8) & 255] << 8) | SBOX[t & 255]
            t ^= RCON[i // nk - 1] << 24
        elif nk > 6 and i % nk == 4:
            t = (SBOX[t >> 24] << 24) | (SBOX[(t >> 16) & 255] << 16) | (SBOX[(t >> 8) & 255] << 8) | SBOX[t & 255]
        w.append(w[i - nk] ^ t)
    return w, rounds


class AES:
    def __init__(self, key):
        w, self.rounds = _expand(key)
        self.ek = [w[4 * i:4 * i + 4] for i in range(self.rounds + 1)]
        r = self.rounds
        # decryption round keys (equivalent inverse cipher): reversed, InvMixColumns applied to the middle ones
        dk = []
        for rnd in range(r, -1, -1):
            k = w[4 * rnd:4 * rnd + 4]
            if 0 < rnd < r:
                k = [D0[SBOX[x >> 24]] ^ D1[SBOX[(x >> 16) & 255]] ^ D2[SBOX[(x >> 8) & 255]] ^ D3[SBOX[x & 255]]
                     for x in k]
            dk.append(k)
        self.dk = dk

    def decrypt_block(self, block):
        dk = self.dk
        s0, s1, s2, s3 = struct.unpack(">IIII", block)
        k = dk[0]
        s0, s1, s2, s3 = s0 ^ k[0], s1 ^ k[1], s2 ^ k[2], s3 ^ k[3]
        for rnd in range(1, self.rounds):
            k = dk[rnd]
            t0 = D0[s0 >> 24] ^ D1[(s3 >> 16) & 255] ^ D2[(s2 >> 8) & 255] ^ D3[s1 & 255] ^ k[0]
            t1 = D0[s1 >> 24] ^ D1[(s0 >> 16) & 255] ^ D2[(s3 >> 8) & 255] ^ D3[s2 & 255] ^ k[1]
            t2 = D0[s2 >> 24] ^ D1[(s1 >> 16) & 255] ^ D2[(s0 >> 8) & 255] ^ D3[s3 & 255] ^ k[2]
            t3 = D0[s3 >> 24] ^ D1[(s2 >> 16) & 255] ^ D2[(s1 >> 8) & 255] ^ D3[s0 & 255] ^ k[3]
            s0, s1, s2, s3 = t0, t1, t2, t3
        k = dk[self.rounds]
        out = bytes([
            INV[s0 >> 24], INV[(s3 >> 16) & 255], INV[(s2 >> 8) & 255], INV[s1 & 255],
            INV[s1 >> 24], INV[(s0 >> 16) & 255], INV[(s3 >> 8) & 255], INV[s2 & 255],
            INV[s2 >> 24], INV[(s1 >> 16) & 255], INV[(s0 >> 8) & 255], INV[s3 & 255],
            INV[s3 >> 24], INV[(s2 >> 16) & 255], INV[(s1 >> 8) & 255], INV[s0 & 255]])
        return bytes(a ^ b for a, b in zip(out, struct.pack(">IIII", *k)))

    def encrypt_block(self, block):
        ek = self.ek
        s0, s1, s2, s3 = struct.unpack(">IIII", block)
        k = ek[0]
        s0, s1, s2, s3 = s0 ^ k[0], s1 ^ k[1], s2 ^ k[2], s3 ^ k[3]
        for rnd in range(1, self.rounds):
            k = ek[rnd]
            t0 = E0[s0 >> 24] ^ E1[(s1 >> 16) & 255] ^ E2[(s2 >> 8) & 255] ^ E3[s3 & 255] ^ k[0]
            t1 = E0[s1 >> 24] ^ E1[(s2 >> 16) & 255] ^ E2[(s3 >> 8) & 255] ^ E3[s0 & 255] ^ k[1]
            t2 = E0[s2 >> 24] ^ E1[(s3 >> 16) & 255] ^ E2[(s0 >> 8) & 255] ^ E3[s1 & 255] ^ k[2]
            t3 = E0[s3 >> 24] ^ E1[(s0 >> 16) & 255] ^ E2[(s1 >> 8) & 255] ^ E3[s2 & 255] ^ k[3]
            s0, s1, s2, s3 = t0, t1, t2, t3
        k = ek[self.rounds]
        out = bytes([
            SBOX[s0 >> 24], SBOX[(s1 >> 16) & 255], SBOX[(s2 >> 8) & 255], SBOX[s3 & 255],
            SBOX[s1 >> 24], SBOX[(s2 >> 16) & 255], SBOX[(s3 >> 8) & 255], SBOX[s0 & 255],
            SBOX[s2 >> 24], SBOX[(s3 >> 16) & 255], SBOX[(s0 >> 8) & 255], SBOX[s1 & 255],
            SBOX[s3 >> 24], SBOX[(s0 >> 16) & 255], SBOX[(s1 >> 8) & 255], SBOX[s2 & 255]])
        return bytes(a ^ b for a, b in zip(out, struct.pack(">IIII", *k)))

    def encrypt_cbc(self, data, iv):
        out, prev = [], iv
        for i in range(0, len(data), 16):
            block = self.encrypt_block(bytes(a ^ b for a, b in zip(data[i:i + 16], prev)))
            out.append(block)
            prev = block
        return b"".join(out)

    def decrypt_ecb(self, data):
        return b"".join(self.decrypt_block(data[i:i + 16]) for i in range(0, len(data) - len(data) % 16, 16))

    def decrypt_cbc(self, data, iv):
        out, prev = [], iv
        for i in range(0, len(data) - len(data) % 16, 16):
            block = data[i:i + 16]
            out.append(bytes(a ^ b for a, b in zip(self.decrypt_block(block), prev)))
            prev = block
        return b"".join(out)


# ---------- OLE compound file ----------

class CompoundFile:
    """Just enough of [MS-CFB] to read streams by name."""

    def __init__(self, data):
        if data[:8] != OLE_MAGIC:
            raise NotEncrypted("not an OLE compound file")
        self.data = data
        self.ssize = 1 << struct.unpack_from("<H", data, 30)[0]
        self.msize = 1 << struct.unpack_from("<H", data, 32)[0]
        n_fat, first_dir = struct.unpack_from("<II", data, 44)
        self.mini_cutoff, first_minifat, n_minifat, first_difat, n_difat = struct.unpack_from("<IIIII", data, 56)
        difat = list(struct.unpack_from("<109I", data, 76))
        sec = first_difat
        for _ in range(n_difat):
            vals = struct.unpack_from(f"<{self.ssize // 4}I", self.sector(sec))
            difat += vals[:-1]
            sec = vals[-1]
        self.fat = []
        for s in difat[:n_fat]:
            self.fat += struct.unpack_from(f"<{self.ssize // 4}I", self.sector(s))
        dirdata = self.chain(first_dir)
        self.entries = []
        for i in range(0, len(dirdata), 128):
            e = dirdata[i:i + 128]
            name_len = struct.unpack_from("<H", e, 64)[0]
            name = e[:max(name_len - 2, 0)].decode("utf-16-le", "replace")
            kind = e[66]
            start, size = struct.unpack_from("<IQ", e, 116)
            self.entries.append((name, kind, start, size))
        root = self.entries[0]
        self.ministream = self.chain(root[2])[:root[3]]
        self.minifat = []
        if n_minifat:
            mf = self.chain(first_minifat)
            self.minifat = list(struct.unpack_from(f"<{len(mf) // 4}I", mf))

    def sector(self, n):
        off = (n + 1) * self.ssize
        return self.data[off:off + self.ssize]

    def chain(self, start):
        out, s, seen = [], start, set()
        while s < 0xFFFFFFFA and s not in seen:
            seen.add(s)
            out.append(self.sector(s))
            s = self.fat[s] if s < len(self.fat) else 0xFFFFFFFE
        return b"".join(out)

    def mini_chain(self, start):
        out, s, seen = [], start, set()
        while s < 0xFFFFFFFA and s not in seen:
            seen.add(s)
            out.append(self.ministream[s * self.msize:(s + 1) * self.msize])
            s = self.minifat[s] if s < len(self.minifat) else 0xFFFFFFFE
        return b"".join(out)

    def stream(self, name):
        for n, kind, start, size in self.entries:
            if n == name and kind == 2:
                data = self.mini_chain(start) if size < self.mini_cutoff else self.chain(start)
                return data[:size]
        raise KeyError(name)


# ---------- ECMA-376 encryption ----------

BLOCK_VERIFIER_INPUT = bytes([0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79])
BLOCK_VERIFIER_VALUE = bytes([0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E])
BLOCK_KEY = bytes([0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6])


def _hash(name):
    return lambda data: hashlib.new(name.replace("-", "").lower(), data).digest()


def _fit(b, n, pad):
    return b[:n] if len(b) >= n else b + bytes([pad]) * (n - len(b))


def _agile(info, package, password):
    root = ET.fromstring(info[8:])
    ns = {"e": "http://schemas.microsoft.com/office/2006/encryption",
          "p": "http://schemas.microsoft.com/office/2006/keyEncryptor/password"}
    kd = root.find("e:keyData", ns)
    ek = root.find(".//p:encryptedKey", ns)
    if kd is None or ek is None:
        raise ValueError("unsupported encryption")
    H = _hash(ek.get("hashAlgorithm"))
    salt = base64.b64decode(ek.get("saltValue"))
    h = H(salt + password.encode("utf-16-le"))
    for i in range(int(ek.get("spinCount"))):
        h = H(struct.pack("<I", i) + h)
    key_bytes = int(ek.get("keyBits")) // 8

    def derived(block):
        return _fit(H(h + block), key_bytes, 0x36)

    def dec(block, value):
        return AES(derived(block)).decrypt_cbc(base64.b64decode(value), salt)
    vin = dec(BLOCK_VERIFIER_INPUT, ek.get("encryptedVerifierHashInput"))
    vval = dec(BLOCK_VERIFIER_VALUE, ek.get("encryptedVerifierHashValue"))
    hsize = int(ek.get("hashSize"))
    if H(vin[:int(ek.get("saltSize"))])[:hsize] != vval[:hsize]:
        raise WrongPassword()
    secret = dec(BLOCK_KEY, ek.get("encryptedKeyValue"))[:int(kd.get("keyBits")) // 8]
    KH = _hash(kd.get("hashAlgorithm"))
    kd_salt = base64.b64decode(kd.get("saltValue"))
    block_size = int(kd.get("blockSize"))
    total = struct.unpack_from("<Q", package, 0)[0]
    aes = AES(secret)
    out = []
    data = package[8:]
    for i in range(0, len(data), 4096):
        iv = _fit(KH(kd_salt + struct.pack("<I", i // 4096)), block_size, 0x36)
        out.append(aes.decrypt_cbc(data[i:i + 4096], iv))
    return b"".join(out)[:total]


def _standard(info, package, password):
    flags, header_size = struct.unpack_from("<II", info, 4)
    header = info[12:12 + header_size]
    key_size = struct.unpack_from("<I", header, 16)[0] or 128
    v = info[12 + header_size:]
    salt_size = struct.unpack_from("<I", v, 0)[0]
    salt = v[4:4 + salt_size]
    enc_verifier = v[4 + salt_size:20 + salt_size]
    enc_hash = v[24 + salt_size:56 + salt_size]
    h = hashlib.sha1(salt + password.encode("utf-16-le")).digest()
    for i in range(50000):
        h = hashlib.sha1(struct.pack("<I", i) + h).digest()
    h = hashlib.sha1(h + struct.pack("<I", 0)).digest()
    x1 = hashlib.sha1(bytes(b ^ 0x36 for b in _fit(h, 64, 0))).digest()
    x2 = hashlib.sha1(bytes(b ^ 0x5C for b in _fit(h, 64, 0))).digest()
    key = (x1 + x2)[:key_size // 8]
    aes = AES(key)
    verifier = aes.decrypt_ecb(enc_verifier)
    if hashlib.sha1(verifier).digest() != aes.decrypt_ecb(enc_hash)[:20]:
        raise WrongPassword()
    total = struct.unpack_from("<Q", package, 0)[0]
    return aes.decrypt_ecb(package[8:])[:total]


def is_encrypted(path):
    """An Excel file saved with a password (an OLE container holding the encrypted workbook)."""
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except OSError:
        return False
    if head != OLE_MAGIC:
        return False
    try:
        with open(path, "rb") as f:
            cf = CompoundFile(f.read())
        cf.stream("EncryptionInfo")
        return True
    except (KeyError, NotEncrypted, struct.error):
        return False


def decrypt(path, password):
    with open(path, "rb") as f:
        cf = CompoundFile(f.read())
    info = cf.stream("EncryptionInfo")
    package = cf.stream("EncryptedPackage")
    major, minor = struct.unpack_from("<HH", info, 0)
    if (major, minor) == (4, 4):
        data = _agile(info, package, password)
    elif minor == 2 and major in (2, 3, 4):
        data = _standard(info, package, password)
    else:
        raise ValueError(f"unsupported encryption version {major}.{minor}")
    if data[:2] != b"PK":
        raise WrongPassword()
    return data


# ---------- writing: OLE compound file ----------

FREE, END, FATSECT, DIFSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFC


def _cfb_key(name):
    """Directory sibling order: shorter names first, then by upper-case name."""
    return (len(name), name.upper())


def write_cfb(tree):
    """tree: {name: bytes (stream) or dict (storage)} -> compound file bytes (version 3, 512-byte sectors)."""
    entries = []          # [name, type, left, right, child, data]

    def add(name, kind, data=None):
        entries.append([name, kind, FREE, FREE, FREE, data])
        return len(entries) - 1

    def build(children):
        """Add a storage's children; returns the id of the root of their (balanced) sibling tree."""
        ids = []
        for n in sorted(children, key=_cfb_key):
            v = children[n]
            if isinstance(v, dict):
                i = add(n, 1)
                entries[i][4] = build(v)
            else:
                i = add(n, 2, v)
            ids.append(i)

        def tree_of(lo, hi):
            if lo > hi:
                return FREE
            mid = (lo + hi) // 2
            entries[ids[mid]][2] = tree_of(lo, mid - 1)
            entries[ids[mid]][3] = tree_of(mid + 1, hi)
            return ids[mid]
        return tree_of(0, len(ids) - 1)

    root = add("Root Entry", 5)
    entries[root][4] = build(tree)
    # small streams (under 4096 bytes) live in the mini stream, in 64-byte mini sectors
    ministream, minifat, start = bytearray(), [], {}
    for i, e in enumerate(entries):
        if e[1] == 2 and len(e[5]) < 4096:
            n = (len(e[5]) + 63) // 64
            start[i] = len(minifat) if n else END
            minifat += [len(minifat) + k + 1 for k in range(n - 1)] + ([END] if n else [])
            ministream += e[5] + b"\0" * (n * 64 - len(e[5]))
    big = [i for i, e in enumerate(entries) if e[1] == 2 and len(e[5]) >= 4096]

    def sec(n):
        return (n + 511) // 512
    dir_secs = sec(len(entries) * 128)
    minifat_secs = sec(len(minifat) * 4)
    mini_secs = sec(len(ministream))
    body = dir_secs + minifat_secs + mini_secs + sum(sec(len(entries[i][5])) for i in big)
    fat_secs = difat_secs = 0
    while True:   # FAT (and DIFAT) sizes depend on the total number of sectors
        need_fat = (fat_secs + difat_secs + body + 127) // 128
        need_difat = max(0, (need_fat - 109 + 126) // 127)
        if (need_fat, need_difat) == (fat_secs, difat_secs):
            break
        fat_secs, difat_secs = need_fat, need_difat
    fat = [FREE] * (fat_secs * 128)
    fat_ids = list(range(fat_secs))
    for s in fat_ids:
        fat[s] = FATSECT
    difat_ids = list(range(fat_secs, fat_secs + difat_secs))
    for s in difat_ids:
        fat[s] = DIFSECT
    pos = fat_secs + difat_secs

    def chain(n):
        nonlocal pos
        if n == 0:
            return END
        first = pos
        for k in range(n):
            fat[pos + k] = pos + k + 1 if k < n - 1 else END
        pos += n
        return first
    dir_start = chain(dir_secs)
    minifat_start = chain(minifat_secs)
    mini_start = chain(mini_secs)
    big_start = {i: chain(sec(len(entries[i][5]))) for i in big}
    out = bytearray(OLE_MAGIC + b"\0" * 16 + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\0" * 6)
    out += struct.pack("<IIIIIIIII", 0, fat_secs, dir_start, 0, 4096,
                       minifat_start, minifat_secs, difat_ids[0] if difat_ids else END, difat_secs)
    difat_head = fat_ids[:109] + [FREE] * (109 - min(len(fat_ids), 109))
    out += struct.pack("<109I", *difat_head)
    for k in range(fat_secs):
        out += struct.pack("<128I", *fat[k * 128:(k + 1) * 128])
    rest = fat_ids[109:]
    for k in range(difat_secs):
        part = rest[k * 127:(k + 1) * 127]
        nxt = difat_ids[k + 1] if k + 1 < difat_secs else END
        out += struct.pack("<128I", *(part + [FREE] * (127 - len(part)) + [nxt]))
    d = bytearray()
    for i, (name, kind, left, right, child, data) in enumerate(entries):
        nm = name.encode("utf-16-le")
        if kind == 5:
            st, size = (mini_start if ministream else END), len(ministream)
        elif kind == 2:
            st, size = (big_start[i] if i in big else start[i]), len(data)
        else:
            st, size = 0, 0
        d += nm + b"\0" * (64 - len(nm)) + struct.pack("<HBB", len(nm) + 2, kind, 1)
        d += struct.pack("<III", left, right, child) + b"\0" * 36 + struct.pack("<IQ", st, size)
    while len(d) < dir_secs * 512:   # unused directory slots
        d += b"\0" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", FREE, FREE, FREE) + b"\0" * 48
    out += d
    mf = struct.pack(f"<{len(minifat)}I", *minifat)
    out += mf + b"\xff" * (minifat_secs * 512 - len(mf))
    out += ministream + b"\0" * (mini_secs * 512 - len(ministream))
    for i in big:
        data = entries[i][5]
        out += data + b"\0" * (sec(len(data)) * 512 - len(data))
    return bytes(out)


# ---------- writing: ECMA-376 agile encryption ----------

BLOCK_INTEGRITY_KEY = bytes([0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6])
BLOCK_INTEGRITY_VALUE = bytes([0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33])
DATASPACES = {   # fixed content Excel expects next to an encrypted package
    "Version": bytes.fromhex("3c0000004d006900630072006f0073006f00660074002e0043006f006e007400610069006e006500720"
                             "02e004400610074006100530070006100630065007300010000000100000001000000"),
    "DataSpaceMap": bytes.fromhex("08000000010000006800000001000000000000002000000045006e0063007200790070007400650"
                                  "064005000610063006b00610067006500320000005300740072006f006e00670045006e00630072"
                                  "0079007000740069006f006e004400610074006100530070006100630065000000"),
    "StrongEncryptionDataSpace": bytes.fromhex("0800000001000000320000005300740072006f006e00670045006e0063007200"
                                               "79007000740069006f006e005400720061006e00730066006f0072006d000000"),
    "Primary": bytes.fromhex("58000000010000004c0000007b00460046003900410033004600300033002d0035003600450046002d"
                             "0034003600310033002d0042004400440035002d003500410034003100430031004400300037003200"
                             "340036007d004e0000004d006900630072006f0073006f00660074002e0043006f006e00740061006900"
                             "6e00650072002e0045006e006300720079007000740069006f006e005400720061006e00730066006f00"
                             "72006d00000001000000010000000100000000000000000000000000000004000000"),
}


def _pad16(b):
    return b + b"\0" * (-len(b) % 16)


def encrypt(data, password, spin=100000):
    """Plain .xlsx bytes -> password-protected file bytes (Agile encryption: AES-256, SHA-512)."""
    import hmac
    import os

    def H(b):
        return hashlib.sha512(b).digest()

    def b64(b):
        return base64.b64encode(b).decode()
    secret, kd_salt, pw_salt = os.urandom(32), os.urandom(16), os.urandom(16)
    aes = AES(secret)
    out = [struct.pack("<Q", len(data))]
    for i in range(0, len(data), 4096):
        out.append(aes.encrypt_cbc(_pad16(data[i:i + 4096]), H(kd_salt + struct.pack("<I", i // 4096))[:16]))
    package = b"".join(out)
    hmac_key = os.urandom(64)
    hmac_value = hmac.new(hmac_key, package, hashlib.sha512).digest()
    enc_hmac_key = aes.encrypt_cbc(hmac_key, H(kd_salt + BLOCK_INTEGRITY_KEY)[:16])
    enc_hmac_value = aes.encrypt_cbc(hmac_value, H(kd_salt + BLOCK_INTEGRITY_VALUE)[:16])
    h = H(pw_salt + password.encode("utf-16-le"))
    for i in range(spin):
        h = H(struct.pack("<I", i) + h)

    def enc(block, value):
        return AES(_fit(H(h + block), 32, 0x36)).encrypt_cbc(_pad16(value), pw_salt)
    verifier = os.urandom(16)
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
           '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" '
           'xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" '
           'xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">'
           '<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" '
           f'cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="{b64(kd_salt)}"/>'
           f'<dataIntegrity encryptedHmacKey="{b64(enc_hmac_key)}" encryptedHmacValue="{b64(enc_hmac_value)}"/>'
           '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">'
           f'<p:encryptedKey spinCount="{spin}" saltSize="16" blockSize="16" keyBits="256" hashSize="64" '
           'cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" '
           f'saltValue="{b64(pw_salt)}" '
           f'encryptedVerifierHashInput="{b64(enc(BLOCK_VERIFIER_INPUT, verifier))}" '
           f'encryptedVerifierHashValue="{b64(enc(BLOCK_VERIFIER_VALUE, H(verifier)))}" '
           f'encryptedKeyValue="{b64(enc(BLOCK_KEY, secret))}"/>'
           '</keyEncryptor></keyEncryptors></encryption>')
    info = struct.pack("<HHI", 4, 4, 0x40) + xml.encode("utf-8")
    return write_cfb({
        "EncryptionInfo": info,
        "EncryptedPackage": package,
        "\x06DataSpaces": {
            "Version": DATASPACES["Version"],
            "DataSpaceMap": DATASPACES["DataSpaceMap"],
            "DataSpaceInfo": {"StrongEncryptionDataSpace": DATASPACES["StrongEncryptionDataSpace"]},
            "TransformInfo": {"StrongEncryptionTransform": {"\x06Primary": DATASPACES["Primary"]}},
        },
    })


def encrypt_file(src, dst, password):
    """Write `src` (a plain .xlsx) to `dst` protected with `password`."""
    import os
    with open(src, "rb") as f:
        data = f.read()
    tmp = dst + ".tmp"
    with open(tmp, "wb") as f:
        f.write(encrypt(data, password))
    os.replace(tmp, dst)
