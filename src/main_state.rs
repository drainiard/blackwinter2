use std::collections::HashSet;
use std::path::Path;
use rand;
use rand::distributions::{IndependentSample, Range};
use ggez::*;
use ggez::event::{Keycode, Mod};
use ggez::graphics::Point2;
use ggez::timer;
use starfield::Starfield;

const PLAYER_SPRITES_COUNT: u8 = 3;

pub struct MainState {
    pos_x: f32,
    pos_y: f32,
    speed: f32,
    moving: HashSet<Keycode>,
    stars: Vec<Starfield>,
    counter: u8,
    font: graphics::Font,
}

impl MainState {
    pub fn new(ctx: &mut Context) -> GameResult<Self> {
        ctx.print_resource_stats();

        let font = graphics::Font::new(ctx, "/TerminusTTF.ttf", 8)?;

        let mut rng = rand::thread_rng();
        let mut stars = Vec::with_capacity(100);
        for _ in 0..100 {
            let star = Starfield {
                pos_x: Range::new(0., 328.).ind_sample(&mut rng),
                pos_y: Range::new(0., 432.).ind_sample(&mut rng),
                speed: Range::new(2. / 3., 8. / 3.).ind_sample(&mut rng),
                color: Range::new(20, 220).ind_sample(&mut rng),
            };
            stars.push(star);
        }

        let s = MainState {
            pos_x: 160.0 - 49. / 2.,
            pos_y: 210.0 - 38. / 2.,
            speed: 4.0,
            moving: HashSet::new(),
            stars: stars,
            counter: 0,
            font: font,
        };
        Ok(s)
    }

    fn draw_background(&mut self, ctx: &mut Context) -> GameResult<()> {
        graphics::set_color(ctx, graphics::Color::from_rgb(255, 255, 255))?;
        graphics::set_background_color(ctx, graphics::Color::from_rgb(0, 5, 10));
        Ok(())
    }

    fn draw_stars(&mut self, ctx: &mut Context) -> GameResult<()> {
        for star in self.stars.iter() {
            let point = Point2::new(star.pos_x, star.pos_y);
            graphics::set_color(ctx, graphics::Color::from_rgb(star.color, star.color, star.color))?;
            graphics::points(ctx, &[point], 1.)?;
        }
        Ok(())
    }

    fn draw_player(&mut self, ctx: &mut Context) -> GameResult<()> {
        graphics::set_color(ctx, graphics::Color::from_rgb(255, 255, 255))?;
        self.counter = (self.counter + 1) % (PLAYER_SPRITES_COUNT * 3);
        let counter: u8 = self.counter / 3;
        let image = graphics::Image::new(ctx, Path::new(&format!("/img0-{}.png", counter + 1)))?;
        graphics::draw(ctx, &image, Point2::new(self.pos_x, self.pos_y), 0.)
    }
}

impl event::EventHandler for MainState {
    fn update(&mut self, _ctx: &mut Context) -> GameResult<()> {
        for keycode in self.moving.iter() {
            match *keycode {
                Keycode::Up => {
                    self.pos_y = (0. as f32).max(self.pos_y - self.speed);
                },
                Keycode::Down => {
                    self.pos_y = (413.0 - 38. as f32).min(self.pos_y + self.speed);
                },
                Keycode::Left => {
                    self.pos_x = (0. as f32).max(self.pos_x - self.speed);
                },
                Keycode::Right => {
                    self.pos_x = (320. - 49. as f32).min(self.pos_x + self.speed);
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

        //timer::yield_now();

        Ok(())
    }

    fn draw(&mut self, ctx: &mut Context) -> GameResult<()> {
        graphics::clear(ctx);

        self.draw_background(ctx)?;
        self.draw_stars(ctx)?;
        self.draw_player(ctx)?;

        let fps = format!("{:.*}", 1, timer::get_fps(ctx));
        let text = graphics::Text::new(ctx, &fps, &self.font)?;
        graphics::draw(ctx, &text, Point2::new(0., 0.), 0.)?;

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

