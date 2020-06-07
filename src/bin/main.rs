extern crate blackwinter2;
extern crate ggez;

use blackwinter2::MainState;
use ggez::{
    conf::{WindowMode, WindowSetup},
    event, ContextBuilder, GameResult,
};
use std::env;
use std::path;

pub fn main() -> GameResult {
    let (width, height) = (320., 413.);

    let mut cb = ContextBuilder::new("black_winter", "ggez")
        .window_setup(WindowSetup::default().title("Black Winter II : Final Assault"))
        .window_mode(WindowMode::default().dimensions(width, height));

    if let Ok(manifest_dir) = env::var("CARGO_MANIFEST_DIR") {
        let mut path = path::PathBuf::from(manifest_dir);
        path.push("resources");
        cb = cb.add_resource_path(path);
    }

    let (ctx, events_loop) = &mut cb.build()?;
    let state = &mut MainState::new(ctx, (width, height))?;

    event::run(ctx, events_loop, state)
}
