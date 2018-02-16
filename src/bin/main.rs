extern crate ggez;
extern crate blackwinter2;

use std::env;
use std::path;
use ggez::*;
use blackwinter2::MainState;


pub fn main() {
    let (width, height) = (320, 413);

    let cb = ContextBuilder::new("black_winter", "ggez")
        .window_setup(conf::WindowSetup::default().title("Black Winter II : Final Assault"))
        .window_mode(conf::WindowMode::default().dimensions(width, height));
    let ctx = &mut cb.build().unwrap();

    if let Ok(manifest_dir) = env::var("CARGO_MANIFEST_DIR") {
        let mut path = path::PathBuf::from(manifest_dir);
        path.push("resources");
        ctx.filesystem.mount(&path, true);
    }

    let state = &mut MainState::new(ctx).unwrap();
    event::run(ctx, state).unwrap();
}
