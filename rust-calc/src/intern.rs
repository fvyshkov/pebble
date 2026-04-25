use std::collections::HashMap;

/// Bidirectional string intern table.
/// All strings are assigned a u32 index at deserialization time.
/// Runtime lookups use only u32 indices.
#[derive(serde::Serialize, serde::Deserialize)]
pub struct Interner {
    to_id: HashMap<String, u32>,
    to_str: Vec<String>,
}

impl Interner {
    pub fn new() -> Self {
        Self {
            to_id: HashMap::default(),
            to_str: Vec::new(),
        }
    }

    /// Intern a string, returning its u32 index.
    /// If already interned, returns the existing index.
    pub fn intern(&mut self, s: &str) -> u32 {
        if let Some(&id) = self.to_id.get(s) {
            return id;
        }
        let id = self.to_str.len() as u32;
        self.to_str.push(s.to_string());
        self.to_id.insert(s.to_string(), id);
        id
    }

    /// Intern a lowercase version of the string.
    pub fn intern_lower(&mut self, s: &str) -> u32 {
        let lower = s.to_lowercase();
        self.intern(&lower)
    }

    /// Get the original string for an interned id.
    pub fn get_str(&self, id: u32) -> &str {
        &self.to_str[id as usize]
    }

    /// Check if a string is already interned.
    pub fn get_id(&self, s: &str) -> Option<u32> {
        self.to_id.get(s).copied()
    }

    pub fn len(&self) -> usize {
        self.to_str.len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_intern_basic() {
        let mut int = Interner::new();
        let a = int.intern("hello");
        let b = int.intern("world");
        let c = int.intern("hello");
        assert_eq!(a, c);
        assert_ne!(a, b);
        assert_eq!(int.get_str(a), "hello");
        assert_eq!(int.get_str(b), "world");
    }

    #[test]
    fn test_intern_lower() {
        let mut int = Interner::new();
        let a = int.intern_lower("Hello");
        let b = int.intern_lower("HELLO");
        assert_eq!(a, b);
        assert_eq!(int.get_str(a), "hello");
    }

    #[test]
    fn test_intern_cyrillic() {
        let mut int = Interner::new();
        let a = int.intern_lower("Доход");
        let b = int.intern_lower("ДОХОД");
        assert_eq!(a, b);
        assert_eq!(int.get_str(a), "доход");
    }
}
