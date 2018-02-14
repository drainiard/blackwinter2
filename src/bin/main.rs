extern crate rand;
extern crate ggez;
extern crate blackwinter2;

use std::collections::HashSet;
use rand::distributions::{IndependentSample, Range};
use ggez::*;
use ggez::event::{Keycode, Mod};
use ggez::graphics::{DrawMode, Point2};
use blackwinter2::Starfield;

struct MainState {
    pos_x: f32,
    pos_y: f32,
    speed: f32,
    moving: HashSet<Keycode>,
    stars: Vec<Starfield>,
}

impl MainState {
    fn new(_ctx: &mut Context) -> GameResult<Self> {
        let mut rng = rand::thread_rng();

        let mut stars = Vec::with_capacity(100);
        for _ in 0..100 {
            let star = Starfield {
                pos_x: Range::new(0., 328.).ind_sample(&mut rng),
                pos_y: Range::new(0., 432.).ind_sample(&mut rng),
                speed: Range::new(0., 8. / 3.).ind_sample(&mut rng) + 2. / 3.,
                color: Range::new(0., 2.).ind_sample(&mut rng) + 1.,
            };
            stars.push(star);
        }

        let s = MainState {
            pos_x: 160.0,
            pos_y: 210.0,
            speed: 8.0,
            moving: HashSet::new(),
            stars: stars,
        };
        Ok(s)
    }

    fn draw_stars(&mut self, ctx: &mut Context) -> GameResult<()> {
        let stars: Vec<_> = self.stars.iter().map(|star| {
            Point2::new(star.pos_x, star.pos_y)
        }).collect();
        graphics::set_color(ctx, graphics::Color::from_rgb(240, 240, 240))?;
        graphics::points(
            ctx,
            &stars,
            1.,
        )
    }
}

impl event::EventHandler for MainState {
    fn update(&mut self, _ctx: &mut Context) -> GameResult<()> {
        for keycode in self.moving.iter() {
            match *keycode {
                Keycode::Up => {
                    self.pos_y = (30. as f32).max(self.pos_y - self.speed);
                },
                Keycode::Down => {
                    self.pos_y = (413.0 - 30. as f32).min(self.pos_y + self.speed);
                },
                Keycode::Left => {
                    self.pos_x = (30. as f32).max(self.pos_x - self.speed);
                },
                Keycode::Right => {
                    self.pos_x = (320. - 30. as f32).min(self.pos_x + self.speed);
                },
                _ => {},
            }
        }

        let mut rng = rand::thread_rng();
        for star in self.stars.iter_mut() {
            star.pos_y += star.speed;
            if star.pos_y > 528. {
                star.pos_y = 0.;
                star.pos_x = Range::new(0., 328.).ind_sample(&mut rng);
                star.speed = Range::new(0., 8. / 3.).ind_sample(&mut rng) + 2. / 3.;
            }
        }

        Ok(())
    }

    fn draw(&mut self, ctx: &mut Context) -> GameResult<()> {
        graphics::clear(ctx);

        self.draw_stars(ctx)?;

        graphics::set_color(ctx, graphics::Color::from_rgb(255, 255, 255))?;
        graphics::set_background_color(ctx, graphics::Color::from_rgb(0, 5, 10));
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
