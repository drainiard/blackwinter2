extern crate ggez;

use std::collections::HashSet;
use ggez::*;
use ggez::event::{Keycode, Mod};
use ggez::graphics::{DrawMode, Point2};

struct MainState {
    pos_x: f32,
    pos_y: f32,
    speed: f32,
    moving: HashSet<Keycode>,
}

impl MainState {
    fn new(_ctx: &mut Context) -> GameResult<Self> {
        let s = MainState {
            pos_x: 160.0,
            pos_y: 210.0,
            speed: 8.0,
            moving: HashSet::new()
        };
        Ok(s)
    }
}

impl event::EventHandler for MainState {
    fn update(&mut self, _ctx: &mut Context) -> GameResult<()> {
        for keycode in self.moving.iter() {
            match *keycode {
                Keycode::Up => {
                    self.pos_y = (30.0 as f32).max(self.pos_y - self.speed);
                },
                Keycode::Down => {
                    self.pos_y = (413.0 - 30 as f32).min(self.pos_y + self.speed);
                },
                Keycode::Left => {
                    self.pos_x = (30.0 as f32).max(self.pos_x - self.speed);
                },
                Keycode::Right => {
                    self.pos_x = (320.0 - 30 as f32).min(self.pos_x + self.speed);
                },
                _ => {},
            }
        }

        //self.pos_x = self.pos_x % 800.0 + 1.0;
        Ok(())
    }

    fn draw(&mut self, ctx: &mut Context) -> GameResult<()> {
        graphics::clear(ctx);
        graphics::circle(
            ctx,
            DrawMode::Fill,
            Point2::new(self.pos_x, self.pos_y),
            30.0,
            0.1,
        )?;
        graphics::present(ctx);
        Ok(())
    }

    fn key_down_event(&mut self, ctx: &mut Context, keycode: Keycode, _keymod: Mod, _repeat: bool) {
        if let Keycode::Escape = keycode {
            let _result = ctx.quit();
            _result.expect("Error on quit");
        }
        self.moving.insert(keycode);
    }

    fn key_up_event(&mut self, _ctx: &mut Context, keycode: Keycode, _keymod: Mod, _repeat: bool) {
        self.moving.remove(&keycode);
    }
}

pub fn main() {
    let (width, height) = (320, 413);
    //let image = "img1-1.bmp";

    let cb = ContextBuilder::new("black_winter", "ggez")
        .window_setup(conf::WindowSetup::default().title("Black Winter II : Final Assault"))
        .window_mode(conf::WindowMode::default().dimensions(width, height));
    let ctx = &mut cb.build().unwrap();
    //let ctx = &mut Context::load_from_conf("super_simple", "ggez", c).unwrap();

    let state = &mut MainState::new(ctx).unwrap();
    event::run(ctx, state).unwrap();
}
