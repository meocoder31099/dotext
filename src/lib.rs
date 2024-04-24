#![allow(unused_imports, dead_code, unused_must_use)]

/**
 * Copyright 2017 Robin Syihab. All rights reserved.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software
 * and associated documentation files (the "Software"), to deal in the Software without restriction,
 * including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
 * and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
 * subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies
 * or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
 * ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
 * IN THE SOFTWARE.
 *
 */
extern crate quick_xml as xml;
extern crate zip;

pub mod document;
pub mod docx;
pub mod odp;
pub mod ods;
pub mod odt;
pub mod pptx;
pub mod xlsx;

pub use document::Document;
pub use document::DocumentKind;
pub use docx::Docx;
pub use odp::Odp;
pub use ods::Ods;
pub use odt::Odt;
pub use pptx::Pptx;
pub use xlsx::Xlsx;

/// This function tries to extract the text from a stream.
/// The filename extension is used to detect the right extraction method.
pub fn extract<R>(reader: R, filename: &str) -> std::io::Result<String>
where
    R: std::io::Read + std::io::Seek,
{
    use std::str::FromStr;

    let extension = filename
        .rsplit_once('.')
        .map(|(_, e)| e)
        .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::Other, "No file extension found"))?;

    DocumentKind::from_str(extension)?.extract(reader)
}

/// This function tries to extract the text from a file.
/// The filename extension is used to detect the right extraction method.
pub fn extract_file<P>(path: P) -> std::io::Result<String>
where
    P: AsRef<std::path::Path>,
{
    let filename = path
        .as_ref()
        .file_name()
        .and_then(|s| s.to_str())
        .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::Other, "No filename found"))?;

    let file = std::fs::File::open(path.as_ref())?;
    extract(file, filename)
}
