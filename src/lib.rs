#![allow(non_upper_case_globals)]
#![allow(non_camel_case_types)]
#![allow(non_snake_case)]
#![allow(dead_code)]

include!("./bindings.rs");

use num::NumCast;
use std::ptr::null_mut;
use widestring::U16CString;

pub struct ExcelBook {
    handle: *mut tagBookHandle,
    is_xlsx: bool,
    active_sheet: *mut tagSheetHandle,
}

pub trait SheetIndex {
    fn sheet_index(&self, book: &ExcelBook) -> i32;
}
impl SheetIndex for i32 {
    fn sheet_index(&self, book: &ExcelBook) -> i32 {
        if *self <= book.sheet_count() && *self >= 0 {
            return *self;
        }
        0
    }
}
impl SheetIndex for &str {
    fn sheet_index(&self, book: &ExcelBook) -> i32 {
        let count = book.sheet_count();
        for i in 0..count {
            if book.sheet_name(i) == *self {
                return i;
            }
        }
        0
    }
}

fn u16ptr_zero_len(p: *const u16) -> usize {
    unsafe {
        let mut len = 0;
        while *p.wrapping_offset(len) != 0 {
            len += 1;
        }
        len as usize
    }
}

impl ExcelBook {
    pub fn new(is_xlsx: bool) -> ExcelBook {
        ExcelBook {
            is_xlsx,
            handle: unsafe {
                if is_xlsx {
                    xlCreateXMLBookCW()
                } else {
                    xlCreateBookCW()
                }
            },
            active_sheet: null_mut(),
        }
    }

    pub fn set_license(&mut self, license_name: &str, license_key: &str) -> &mut Self {
        let license_name = U16CString::from_str(license_name).unwrap();
        let license_key = U16CString::from_str(license_key).unwrap();
        unsafe {
            xlBookSetKeyW(self.handle, license_name.as_ptr(), license_key.as_ptr());
        }
        self
    }

    pub fn add_sheet(&mut self, name: &str) -> &mut Self {
        let name = U16CString::from_str(name).unwrap();
        self.active_sheet = unsafe { xlBookAddSheetW(self.handle, name.as_ptr(), null_mut()) };
        self
    }
    pub fn select_sheet<T: SheetIndex>(&mut self, index: T) -> &mut Self {
        let index = index.sheet_index(self);
        self.active_sheet = unsafe { xlBookGetSheetW(self.handle, index) };
        self
    }
    pub fn sheet_count(&self) -> i32 {
        unsafe { xlBookSheetCountW(self.handle) }
    }
    pub fn sheet_name(&self, index: i32) -> String {
        unsafe {
            let name = xlBookGetSheetNameW(self.handle, index);
            let name = U16CString::from_ptr(name, u16ptr_zero_len(name)).unwrap();
            name.to_string().unwrap()
        }
    }

    pub fn read_str(&self, row: i32, col: i32) -> String {
        unsafe {
            let value = xlSheetReadStrW(self.active_sheet, row, col, null_mut());
            let value = U16CString::from_ptr(value, u16ptr_zero_len(value)).unwrap();
            value.to_string().unwrap()
        }
    }

    pub fn write_str<T: AsRef<str>>(&mut self, row: i32, col: i32, value: T) -> &mut Self {
        let value = U16CString::from_str(value.as_ref()).unwrap();
        unsafe {
            xlSheetWriteStrW(self.active_sheet, row, col, value.as_ptr(), null_mut());
        }
        self
    }

    pub fn read_num(&self, row: i32, col: i32) -> f64 {
        unsafe { xlSheetReadNumW(self.active_sheet, row, col, null_mut()) }
    }

    pub fn write_num<T: NumCast>(&mut self, row: i32, col: i32, value: T) -> &mut Self {
        let value: f64 = value.to_f64().unwrap();
        unsafe {
            xlSheetWriteNumW(self.active_sheet, row, col, value, null_mut());
        }
        self
    }

    pub fn read_bool(&self, row: i32, col: i32) -> bool {
        let value = unsafe { xlSheetReadBoolW(self.active_sheet, row, col, null_mut()) };
        if value == 0 {
            false
        } else {
            true
        }
    }

    pub fn write_bool(&mut self, row: i32, col: i32, value: bool) -> &mut Self {
        let value: i32 = if value { 1 } else { 0 };
        unsafe {
            xlSheetWriteBoolW(self.active_sheet, row, col, value, null_mut());
        }
        self
    }
    pub fn load(&mut self, path: &str) -> &mut Self {
        let path = U16CString::from_str(path).unwrap();
        unsafe {
            xlBookLoadW(self.handle, path.as_ptr());
        }
        self
    }
    pub fn save(&self, path: &str) {
        let path = U16CString::from_str(path).unwrap();
        unsafe {
            xlBookSaveW(self.handle, path.as_ptr());
        }
    }
    pub fn version(&self) -> String {
        let version: i32;
        unsafe {
            version = xlBookVersionW(self.handle);
        }
        format!("{:x}", version)
    }
}

impl Drop for ExcelBook {
    fn drop(&mut self) {
        unsafe {
            xlBookReleaseW(self.handle);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_excel() {
        let mut book = ExcelBook::new(true);
        print!("Version: {}\n", book.version());
        book.add_sheet("Sheet1")
            .write_str(1, 0, "This is a Excel Book中文")
            .write_num(2, 0, 1234567890)
            .write_bool(3, 0, true)
            .save("test.xlsx");
        let name = book.sheet_name(0);
        assert_eq!(name, "Sheet1");
        let count = book.sheet_count();
        assert_eq!(count, 1);
        let str = book.read_str(1, 0);
        assert_eq!(str, "This is a Excel Book中文");
        let num = book.read_num(2, 0);
        assert_eq!(num, 1234567890.0);
        let bool = book.read_bool(3, 0);
        assert_eq!(bool, true);
        println!("{} {} {} {} {}", name, count, str, num, bool);
    }
}
