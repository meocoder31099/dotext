use zip::ZipArchive;

use xml::events::Event;
use xml::reader::Reader;

use std::clone::Clone;
use std::fs::File;
use std::io;
use std::io::prelude::*;
use std::path::{Path, PathBuf};
use std::str::FromStr;
use zip::read::ZipFile;

use crate::{Docx, Odp, Ods, Odt, Pptx, Xlsx};

pub enum DocumentKind {
    Docx,
    Odp,
    Ods,
    Odt,
    Pptx,
    Xlsx,
}

impl DocumentKind {
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Docx => "Word Document",
            Self::Odp => "Open Office Presentation",
            Self::Ods => "Open Office Spreadsheet",
            Self::Odt => "Open Office Document",
            Self::Pptx => "Power Point",
            Self::Xlsx => "Excel",
        }
    }

    pub fn extension(&self) -> &'static str {
        match self {
            Self::Docx => "docx",
            Self::Odp => "odp",
            Self::Ods => "ods",
            Self::Odt => "Odt",
            Self::Pptx => "pptx",
            Self::Xlsx => "xlsx",
        }
    }

    /// Read the document from a reader, like a buffer
    pub fn extract<R>(&self, reader: R) -> io::Result<String>
    where
        R: Read + io::Seek,
    {
        let mut isi = String::new();

        match self {
            DocumentKind::Docx => Docx::from_reader(reader)?.read_to_string(&mut isi),
            DocumentKind::Odp => Odp::from_reader(reader)?.read_to_string(&mut isi),
            DocumentKind::Ods => Ods::from_reader(reader)?.read_to_string(&mut isi),
            DocumentKind::Odt => Odt::from_reader(reader)?.read_to_string(&mut isi),
            DocumentKind::Pptx => Pptx::from_reader(reader)?.read_to_string(&mut isi),
            DocumentKind::Xlsx => Xlsx::from_reader(reader)?.read_to_string(&mut isi),
        };

        Ok(isi)
    }
}

impl FromStr for DocumentKind {
    type Err = io::Error;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "docx" => Ok(Self::Docx),
            "odp" => Ok(Self::Odp),
            "ods" => Ok(Self::Ods),
            "Odt" => Ok(Self::Odt),
            "pptx" => Ok(Self::Pptx),
            "xlsx" => Ok(Self::Xlsx),
            _ => Err(io::Error::new(
                io::ErrorKind::Other,
                "File format not supported",
            )),
        }
    }
}

pub trait Document<T>: Read {
    /// Returns the document type
    fn kind(&self) -> DocumentKind;

    /// Read the document from the disk
    fn open<P>(path: P) -> io::Result<T>
    where
        P: AsRef<Path>,
    {
        let file = File::open(path.as_ref())?;
        Self::from_reader(file)
    }

    /// Read the document from a reader, like a buffer
    fn from_reader<R>(reader: R) -> io::Result<T>
    where
        R: Read + io::Seek;
}

pub(crate) fn open_doc_read_data<R>(
    reader: R,
    content_name: &str,
    tags: &[&str],
) -> io::Result<String>
where
    R: Read + io::Seek,
{
    let mut archive = ZipArchive::new(reader)?;

    let mut xml_data = String::new();

    for i in 0..archive.len() {
        let mut c_file = archive.by_index(i).unwrap();
        if c_file.name() == content_name {
            c_file.read_to_string(&mut xml_data);
            break;
        }
    }

    let mut xml_reader = Reader::from_str(xml_data.as_ref());

    let mut buf = Vec::new();
    let mut txt = Vec::new();

    if xml_data.len() > 0 {
        let mut to_read = false;
        loop {
            match xml_reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    for tag in tags {
                        if e.name().as_ref() == tag.as_bytes() {
                            to_read = true;
                            if e.name().as_ref() == b"text:p" {
                                txt.push("\n\n".to_string());
                            }
                            break;
                        }
                    }
                }
                Ok(Event::Text(e)) => {
                    if to_read {
                        txt.push(e.unescape().unwrap().into_owned());
                        to_read = false;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(io::Error::new(
                        io::ErrorKind::Other,
                        format!(
                            "Error at position {}: {:?}",
                            xml_reader.buffer_position(),
                            e
                        ),
                    ))
                }
                _ => (),
            }
        }
    }

    Ok(txt.join(""))
}
