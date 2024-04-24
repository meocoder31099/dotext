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

pub struct Odp {
    data: Cursor<String>,
}

impl Document<Odp> for Odp {
    fn kind(&self) -> DocumentKind {
        DocumentKind::Odp
    }

    fn from_reader<R>(reader: R) -> io::Result<Odp>
    where
        R: Read + io::Seek,
    {
        let text = open_doc_read_data(reader, "content.xml", &["text:p", "text:span"])?;

        Ok(Odp {
            data: Cursor::new(text),
        })
    }
}

impl Read for Odp {
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
        let _ = Odp::open(Path::new("samples/sample.odp"));
    }

    #[test]
    fn read() {
        let mut f = Odp::open(Path::new("samples/sample.odp")).unwrap();

        let mut data = String::new();
        let len = f.read_to_string(&mut data).unwrap();
        println!("len: {}, data: {}", len, data);
    }
}
