use std::hash::{Hash, Hasher};

/// Maximum number of analytics (axes) per sheet.
pub const MAX_AXES: usize = 8;

/// A cell coordinate encoded as a fixed-size array of interned record IDs.
/// Copy type — zero heap allocation on pass/return.
#[derive(Clone, Copy, Eq, serde::Serialize, serde::Deserialize)]
pub struct CoordKey {
    pub rids: [u32; MAX_AXES],
    pub len: u8,
}

impl CoordKey {
    /// Create from a slice of interned record IDs.
    pub fn new(rids: &[u32]) -> Self {
        debug_assert!(rids.len() <= MAX_AXES);
        let mut key = CoordKey {
            rids: [0; MAX_AXES],
            len: rids.len() as u8,
        };
        key.rids[..rids.len()].copy_from_slice(rids);
        key
    }

    /// Return a new CoordKey with one axis replaced. Zero allocation.
    #[inline]
    pub fn with_axis(&self, axis_idx: usize, new_rid: u32) -> Self {
        let mut copy = *self;
        copy.rids[axis_idx] = new_rid;
        copy
    }

    /// Get the record ID at a given axis index.
    #[inline]
    pub fn get(&self, axis_idx: usize) -> u32 {
        self.rids[axis_idx]
    }
}

impl PartialEq for CoordKey {
    #[inline]
    fn eq(&self, other: &Self) -> bool {
        if self.len != other.len {
            return false;
        }
        let n = self.len as usize;
        self.rids[..n] == other.rids[..n]
    }
}

impl Hash for CoordKey {
    #[inline]
    fn hash<H: Hasher>(&self, state: &mut H) {
        let n = self.len as usize;
        self.rids[..n].hash(state);
    }
}

impl std::fmt::Debug for CoordKey {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let n = self.len as usize;
        write!(f, "CK[")?;
        for i in 0..n {
            if i > 0 {
                write!(f, "|")?;
            }
            write!(f, "{}", self.rids[i])?;
        }
        write!(f, "]")
    }
}

/// Cell flags stored as a single u8.
pub const FLAG_COMPUTED: u8 = 0x01;
pub const FLAG_COMPUTING: u8 = 0x02;
pub const FLAG_MANUAL: u8 = 0x04;
pub const FLAG_PHANTOM: u8 = 0x08;
pub const FLAG_UNRESOLVED: u8 = 0x10;
pub const FLAG_HAS_FORMULA: u8 = 0x20;
pub const FLAG_EMPTY_ORIGINAL: u8 = 0x40; // original value was empty/unparseable string

/// Per-cell state with inline original value and flags.
/// No separate HashSets needed.
#[derive(Clone, Copy, serde::Serialize, serde::Deserialize)]
pub struct CellState {
    pub value: f64,
    pub original_value: f64,
    pub formula_id: u32, // u32::MAX = no formula
    pub flags: u8,
}

impl CellState {
    pub const NO_FORMULA: u32 = u32::MAX;

    pub fn new_manual(value: f64) -> Self {
        CellState {
            value,
            original_value: value,
            formula_id: Self::NO_FORMULA,
            flags: FLAG_MANUAL,
        }
    }

    pub fn new_formula(value: f64, formula_id: u32) -> Self {
        CellState {
            value,
            original_value: value,
            formula_id,
            flags: FLAG_HAS_FORMULA,
        }
    }

    pub fn new_phantom(value: f64) -> Self {
        CellState {
            value,
            original_value: value,
            formula_id: Self::NO_FORMULA,
            flags: FLAG_PHANTOM,
        }
    }

    #[inline]
    pub fn has_formula(&self) -> bool {
        self.formula_id != Self::NO_FORMULA
    }

    #[inline]
    pub fn is_computed(&self) -> bool {
        self.flags & FLAG_COMPUTED != 0
    }

    #[inline]
    pub fn is_computing(&self) -> bool {
        self.flags & FLAG_COMPUTING != 0
    }

    #[inline]
    pub fn is_manual(&self) -> bool {
        self.flags & FLAG_MANUAL != 0
    }

    #[inline]
    pub fn is_phantom(&self) -> bool {
        self.flags & FLAG_PHANTOM != 0
    }
}

/// Type alias for the cell key: (interned_sheet_id, CoordKey).
pub type CellKey = (u32, CoordKey);

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_coord_key_basic() {
        let ck = CoordKey::new(&[10, 20, 30]);
        assert_eq!(ck.get(0), 10);
        assert_eq!(ck.get(1), 20);
        assert_eq!(ck.get(2), 30);
        assert_eq!(ck.len, 3);
    }

    #[test]
    fn test_coord_key_with_axis() {
        let ck = CoordKey::new(&[10, 20, 30]);
        let ck2 = ck.with_axis(1, 99);
        assert_eq!(ck2.get(0), 10);
        assert_eq!(ck2.get(1), 99);
        assert_eq!(ck2.get(2), 30);
        // Original unchanged
        assert_eq!(ck.get(1), 20);
    }

    #[test]
    fn test_coord_key_eq_hash() {
        use std::collections::HashMap;
        let ck1 = CoordKey::new(&[1, 2, 3]);
        let ck2 = CoordKey::new(&[1, 2, 3]);
        let ck3 = CoordKey::new(&[1, 2, 4]);
        assert_eq!(ck1, ck2);
        assert_ne!(ck1, ck3);

        let mut map = HashMap::new();
        map.insert(ck1, 42);
        assert_eq!(map.get(&ck2), Some(&42));
        assert_eq!(map.get(&ck3), None);
    }

    #[test]
    fn test_coord_key_copy() {
        let ck = CoordKey::new(&[5, 10]);
        let ck2 = ck; // Copy, not move
        assert_eq!(ck, ck2);
    }

    #[test]
    fn test_cell_state() {
        let mut cell = CellState::new_formula(100.0, 5);
        assert!(cell.has_formula());
        assert!(!cell.is_computed());
        assert!(!cell.is_manual());

        cell.flags |= FLAG_COMPUTED;
        assert!(cell.is_computed());

        let manual = CellState::new_manual(42.0);
        assert!(manual.is_manual());
        assert!(!manual.has_formula());
    }
}
