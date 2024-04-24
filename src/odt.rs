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

pub struct Odt {
    data: Cursor<String>,
}

impl Document<Odt> for Odt {
    fn kind(&self) -> DocumentKind {
        DocumentKind::Odt
    }

    fn from_reader<R>(reader: R) -> io::Result<Odt>
    where
        R: Read + io::Seek,
    {
        let text = open_doc_read_data(reader, "content.xml", &["text:p"])?;

        Ok(Odt {
            data: Cursor::new(text),
        })
    }
}

impl Read for Odt {
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
        let _ = Odt::open(Path::new("samples/sample.odt"));
    }

    #[test]
    fn read() {
        let mut f = Odt::open(Path::new("samples/sample.odt")).unwrap();

        let mut data = String::new();
        let len = f.read_to_string(&mut data).unwrap();
        println!("len: {}, data: {}", len, data);
    }
}
