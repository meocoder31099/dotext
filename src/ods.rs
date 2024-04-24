use zip::ZipArchive;

use xml::events::Event;
use xml::reader::Reader;

use std::clone::Clone;
use std::fs::File;
use std::io;
use std::io::prelude::*;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use zip::read::ZipFile;

use crate::document::{open_doc_read_data, Document, DocumentKind};

pub struct Ods {
    data: Cursor<String>,
}

impl Document<Ods> for Ods {
    fn kind(&self) -> DocumentKind {
        DocumentKind::Ods
    }

    fn from_reader<R>(reader: R) -> io::Result<Ods>
    where
        R: Read + io::Seek,
    {
        let text = open_doc_read_data(reader, "content.xml", &["text:p"])?;

        Ok(Ods {
            data: Cursor::new(text),
        })
    }
}

impl Read for Ods {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        self.data.read(buf)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::{Path, PathBuf};

    #[test]
    fn instantiate() {
        let _ = Ods::open(Path::new("samples/sample.ods"));
    }

    #[test]
    fn read() {
        let mut f = Ods::open(Path::new("samples/sample.ods")).unwrap();

        let mut data = String::new();
        let len = f.read_to_string(&mut data).unwrap();
        println!("len: {}, data: {}", len, data);
    }
}
