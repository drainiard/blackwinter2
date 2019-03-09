#[macro_use]
extern crate nom;

use nom::le_u32;
use std::env;
use std::error::Error;
use std::fs::File;
use std::io::Read;
use std::io::Write;
use std::path;

#[derive(Debug, Eq, PartialEq)]
pub enum FrxExt {
    BMP,
    JPG,
    GIF,
    TIF,
    PNG,
    WMF,
    DAT,
}

#[derive(Debug)]
pub struct FrxItem {
    pub ext: FrxExt,
    pub length: u32,
    pub content: Vec<u8>,
}

named!(frx_item<&[u8], FrxItem>,
    alt!(
        frx_item_bmp // |
        // frx_item_jpg |
        // frx_item_gif |
        // frx_item_png |
        // frx_item_dat
    )
);

named!(frx_item_bmp<&[u8], FrxItem>,
    do_parse!(
        take!(24) >>
        length: le_u32 >>
        peek!(tag!(&[0x42, 0x4D])) >>
        content: take!(length) >>
        (FrxItem { ext: FrxExt::BMP, length: length, content: content.to_vec() })
    )
);

/*
const JPG_BOF: &[u8] = &[0xFF, 0xD8];
const JPG_EOF: &[u8] = &[0xFF, 0xD9];

named!(frx_item_jpg<&[u8], FrxItem>,
    do_parse!(
        tag!(JPG_BOF) >>
        content: take_until_and_consume!(JPG_EOF) >>
        (FrxIt
        em { ext: FrxExt::JPG, content: content.to_vec() })
    )
);
*/

/*
named!(frx_item_gif<&[u8], FrxItem>,
    do_parse!(
        ext: value!(FrxExt::GIF, tag!("GIF")) >>
        alt!(
            tag!("87a") |
            tag!("89a")
        ) >>
        length: le_u32 >>
        content: take!(length) >>
        (FrxItem { ext, content: content.to_vec() })
    )
);

//        value!(FrxExt::PNG, tag!(&[0x89, 0x50])) |

named!(frx_item_dat<&[u8], FrxItem>,
    do_parse!(
        length: le_u32 >>
        content: take!(length) >>
        (FrxItem { ext: FrxExt::DAT, content: content.to_vec() })
    )
);
*/

named!(frx_items<&[u8], Vec<FrxItem>>, many0!(frx_item));

pub fn main() {
    if let Ok(manifest_dir) = env::var("CARGO_MANIFEST_DIR") {
        let mut resources_path = path::PathBuf::from(manifest_dir);
        resources_path.push("resources");
        let mut file = File::open(resources_path.join("FormBW.frx")).expect("file not found");
        let mut buffer: Vec<u8> = Vec::new();
        let _ = file.read_to_end(&mut buffer);

        let mut i = 1;
        loop {
            let tmp_buffer = buffer.clone();
            let (new_buffer, item) = frx_item_bmp(&tmp_buffer)
                .map_err(|err| err.description().to_owned())
                .unwrap();
            if item.ext == FrxExt::BMP {
                println!("{:x?} {:x?}", item.length, item.content.len());
                let mut extracted_file = File::create(
                    resources_path
                        .join("extracted")
                        .join(format!("img_{:04}.bmp", i)),
                )
                .unwrap();
                let _result = extracted_file.write_all(&item.content);
            }
            buffer = new_buffer.to_vec();

            if new_buffer.len() < 1 {
                break;
            }
            i += 1;
        }
    }
}
