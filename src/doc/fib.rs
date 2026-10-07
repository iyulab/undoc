//! The File Information Block — the index at the head of the `WordDocument` stream ([MS-DOC]
//! 2.5.1) that says where everything else is.

use crate::error::{Error, Result};

/// `wIdent` of a Word binary document.
const WORD_IDENT: u16 = 0xA5EC;

/// `wIdent` of a Word 6.0 or Word 95 document.
const WORD6_IDENT: u16 = 0xA5DC;

/// The lowest `nFib` of the Word 97 file format. Word 6.0 and Word 95 wrote 101–104 and use a
/// different, incompatible FIB and text layout.
const FIRST_WORD97_NFIB: u16 = 0x00C0;

/// `FibBase` flag: the document is encrypted or XOR-obfuscated.
const F_ENCRYPTED: u16 = 0x0100;
/// `FibBase` flag: the table stream is `1Table` rather than `0Table`.
const F_WHICH_TBL_STM: u16 = 0x0200;

/// Size of `FibBase`.
const FIB_BASE_LEN: usize = 32;

/// Positions within `FibRgFcLcb97`, counted in (fc, lcb) pairs.
const PAIR_STSHF: usize = 1;
const PAIR_PLCFFND_REF: usize = 2;
const PAIR_PLCFFND_TXT: usize = 3;
const PAIR_PLCF_HDD: usize = 11;
const PAIR_PLCF_BTE_CHPX: usize = 12;
const PAIR_PLCF_BTE_PAPX: usize = 13;
const PAIR_STTBF_FFN: usize = 15;
const PAIR_CLX: usize = 33;
const PAIR_PLCFEND_REF: usize = 46;
const PAIR_PLCFEND_TXT: usize = 47;
const PAIR_PLF_LST: usize = 73;
const PAIR_PLF_LFO: usize = 74;

/// Positions within `FibRgLw97`, counted in 32-bit words.
const LW_CCP_TEXT: usize = 3;
const LW_CCP_FTN: usize = 4;
const LW_CCP_HDD: usize = 5;
const LW_CCP_ATN: usize = 7;

/// A byte range of the table stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct FcLcb {
    pub fc: u32,
    pub lcb: u32,
}

/// What the reader needs from the FIB.
#[derive(Debug, Clone)]
pub(super) struct Fib {
    /// `1Table` or `0Table`.
    pub table_stream: &'static str,
    /// Character count of the main document text — the first story.
    pub ccp_text: u32,
    /// Character counts of the footnote, header and comment stories that follow it — the
    /// endnote story begins after all three.
    pub ccp_ftn: u32,
    pub ccp_hdd: u32,
    pub ccp_atn: u32,
    /// The piece table (`Clx`).
    pub clx: FcLcb,
    /// Paragraph property page index (`PlcBtePapx`).
    pub plcf_bte_papx: FcLcb,
    /// Character property page index (`PlcBteChpx`).
    pub plcf_bte_chpx: FcLcb,
    /// Footnote and endnote reference positions and text ranges.
    pub plcffnd_ref: FcLcb,
    pub plcffnd_txt: FcLcb,
    pub plcfend_ref: FcLcb,
    pub plcfend_txt: FcLcb,
    /// The style sheet (`STSH`).
    pub stshf: FcLcb,
    /// The font table (`SttbfFfn`).
    pub sttbf_ffn: FcLcb,
    /// Header and footer story ranges (`PlcfHdd`).
    pub plcf_hdd: FcLcb,
    /// List definitions and list format overrides.
    pub plf_lst: FcLcb,
    pub plf_lfo: FcLcb,
}

fn u16_at(data: &[u8], at: usize) -> Option<u16> {
    data.get(at..at + 2)
        .map(|b| u16::from_le_bytes([b[0], b[1]]))
}

fn u32_at(data: &[u8], at: usize) -> Option<u32> {
    data.get(at..at + 4)
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

fn word6() -> Error {
    Error::UnsupportedFormat(
        "Word 6.0/95 document (.doc) — only Word 97 and later are supported".into(),
    )
}

fn truncated() -> Error {
    Error::InvalidData("Word document: the File Information Block is truncated".into())
}

/// Read the FIB from the start of the `WordDocument` stream.
pub(super) fn parse(word_document: &[u8]) -> Result<Fib> {
    let data = word_document;
    let ident = u16_at(data, 0).ok_or_else(truncated)?;
    if ident == WORD6_IDENT {
        return Err(word6());
    }
    if ident != WORD_IDENT {
        return Err(Error::InvalidData(format!(
            "Word document: the WordDocument stream does not begin with a File Information \
             Block (wIdent 0x{ident:04X}, expected 0x{WORD_IDENT:04X})"
        )));
    }
    let n_fib = u16_at(data, 2).ok_or_else(truncated)?;
    if n_fib < FIRST_WORD97_NFIB {
        return Err(word6());
    }
    let flags = u16_at(data, 0x0A).ok_or_else(truncated)?;
    if flags & F_ENCRYPTED != 0 {
        return Err(Error::Encrypted);
    }
    let table_stream = if flags & F_WHICH_TBL_STM != 0 {
        "1Table"
    } else {
        "0Table"
    };

    // FibBase, then three variable-length arrays, each preceded by its element count.
    let mut at = FIB_BASE_LEN;
    let csw = u16_at(data, at).ok_or_else(truncated)? as usize;
    at += 2 + csw * 2;
    let cslw = u16_at(data, at).ok_or_else(truncated)? as usize;
    let rg_lw = at + 2;
    at = rg_lw + cslw * 4;
    let cb_rg_fc_lcb = u16_at(data, at).ok_or_else(truncated)? as usize;
    let rg_fc_lcb = at + 2;

    if cslw <= LW_CCP_ATN || cb_rg_fc_lcb <= PAIR_PLF_LFO {
        return Err(truncated());
    }
    let lw = |index: usize| u32_at(data, rg_lw + index * 4).ok_or_else(truncated);
    let pair = |index: usize| -> Result<FcLcb> {
        let at = rg_fc_lcb + index * 8;
        Ok(FcLcb {
            fc: u32_at(data, at).ok_or_else(truncated)?,
            lcb: u32_at(data, at + 4).ok_or_else(truncated)?,
        })
    };

    Ok(Fib {
        table_stream,
        ccp_text: lw(LW_CCP_TEXT)?,
        ccp_ftn: lw(LW_CCP_FTN)?,
        ccp_hdd: lw(LW_CCP_HDD)?,
        ccp_atn: lw(LW_CCP_ATN)?,
        clx: pair(PAIR_CLX)?,
        plcf_bte_papx: pair(PAIR_PLCF_BTE_PAPX)?,
        plcf_bte_chpx: pair(PAIR_PLCF_BTE_CHPX)?,
        plcffnd_ref: pair(PAIR_PLCFFND_REF)?,
        plcffnd_txt: pair(PAIR_PLCFFND_TXT)?,
        plcfend_ref: pair(PAIR_PLCFEND_REF)?,
        plcfend_txt: pair(PAIR_PLCFEND_TXT)?,
        stshf: pair(PAIR_STSHF)?,
        sttbf_ffn: pair(PAIR_STTBF_FFN)?,
        plcf_hdd: pair(PAIR_PLCF_HDD)?,
        plf_lst: pair(PAIR_PLF_LST)?,
        plf_lfo: pair(PAIR_PLF_LFO)?,
    })
}
