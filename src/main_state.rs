use std::collections::HashSet;
use std::path::Path;
use rand;
use rand::distributions::{IndependentSample, Range};
use ggez::*;
use ggez::event::{Keycode, Mod};
use ggez::graphics::Point2;
use ggez::timer;
use starfield::Starfield;
use bullet::{Bullet,Vector2D};

const PLAYER_SPRITES_COUNT: u8 = 3;

pub struct RectSize {
    width: f32,
    height: f32,
}

pub struct MainState {
    font: graphics::Font,
    music: audio::Source,
    window: RectSize,
    pos: Point2,
    speed: f32,
    moving: HashSet<Keycode>,
    stars: Vec<Starfield>,
    counter: u8,
    bullets: Vec<Bullet>,
}

impl MainState {
    pub fn new(ctx: &mut Context) -> GameResult<Self> {
        ctx.print_resource_stats();

        let (width, height) = (ctx.conf.window_mode.width as f32, ctx.conf.window_mode.height as f32);
        let font = graphics::Font::new(ctx, "/TerminusTTF.ttf", 8)?;

        let mut rng = rand::thread_rng();
        let mut stars = Vec::with_capacity(100);
        for _ in 0..100 {
            let star = Starfield {
                pos: Point2::new(Range::new(0., width).ind_sample(&mut rng), Range::new(0., height).ind_sample(&mut rng)),
                speed: Range::new(0.5 / 3., 4. / 3.).ind_sample(&mut rng),
                color: Range::new(0, 160).ind_sample(&mut rng),
            };
            stars.push(star);
        }

        let music = audio::Source::new(ctx, Path::new("/bgplay.ogg"))?;
        println!("{:?}", music.play());


        let s = MainState {
            font: font,
            music: music,
            window: RectSize { width: width, height: height },
            pos: Point2::new((width - 49.) / 2., (height - 38.) / 2.),
            speed: 3.0,
            moving: HashSet::new(),
            stars: stars,
            counter: 0,
            bullets: Vec::new(),
        };
        Ok(s)
    }

    fn shoot_main(&mut self, ctx: &mut Context) -> GameResult<()> {
        // TODO add logic to change behavior based on configured gun

        let image = graphics::Image::new(ctx, Path::new("/img0-1.png"))?;
        let bullet_image = graphics::Image::new(ctx, Path::new("/img6-1.png"))?;
        let adjust = Vector2D::new((image.width() - bullet_image.width()) as f32 / 2., bullet_image.height() as f32 * -1.);

        self.bullets.push(Bullet::new(self.pos.x + adjust.x, self.pos.y + adjust.y, 0., self.speed * -6.));

        Ok(())
    }

    fn draw_background(&mut self, ctx: &mut Context) -> GameResult<()> {
        graphics::set_color(ctx, graphics::Color::from_rgb(255, 255, 255))?;
        graphics::set_background_color(ctx, graphics::Color::from_rgb(0, 5, 10));
        Ok(())
    }

    fn draw_stars(&mut self, ctx: &mut Context) -> GameResult<()> {
        for star in self.stars.iter() {
            let point = Point2::new(star.pos.x, star.pos.y);
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
        graphics::draw(ctx, &image, Point2::new(self.pos.x, self.pos.y), 0.)?;

        let bullet_image = graphics::Image::new(ctx, Path::new(&format!("/img6-1.png")))?;

        let dead_bullets: Vec<usize> = self.bullets.iter().enumerate().filter(|&(_i, bullet): &(usize, &Bullet)| {
            bullet.position.y < -413.
        }).map(|(i, _)| i).rev().collect();
        for i in dead_bullets {
            self.bullets.swap_remove(i);
        }

        for bullet in self.bullets.iter() {
            let point = bullet.position.to_point2();
            graphics::draw(ctx, &bullet_image, point, 0.)?;
        }

        Ok(())
    }
}

impl event::EventHandler for MainState {
    fn update(&mut self, ctx: &mut Context) -> GameResult<()> {
        for bullet in self.bullets.iter_mut() {
            bullet.update(ctx)?;
        }

        let speed = self.speed;
        for keycode in self.moving.clone() {
            match keycode {
                Keycode::Up => {
                    let new_pos_y = (0. as f32).max(self.pos.y - speed);
                    let diff_pos_y = self.pos.y - new_pos_y;
                    self.pos.y = new_pos_y;
                    for bullet in self.bullets.iter_mut() {
                        bullet.position.y = bullet.position.y - diff_pos_y;
                    }
                },
                Keycode::Down => {
                    let new_pos_y = (self.window.height - 38. as f32).min(self.pos.y + self.speed);
                    let diff_pos_y = new_pos_y - self.pos.y;
                    self.pos.y = new_pos_y;
                    for bullet in self.bullets.iter_mut() {
                        bullet.position.y = bullet.position.y + diff_pos_y;
                    }
                },
                Keycode::Left => {
                    let new_pos_x = (0. as f32).max(self.pos.x - self.speed);
                    let diff_pos_x = self.pos.x - new_pos_x;
                    self.pos.x = new_pos_x;
                    for bullet in self.bullets.iter_mut() {
                        bullet.position.x = bullet.position.x - diff_pos_x;
                    }
                },
                Keycode::Right => {
                    let new_pos_x = (self.window.width - 49. as f32).min(self.pos.x + self.speed);
                    let diff_pos_x = new_pos_x - self.pos.x;
                    self.pos.x = new_pos_x;
                    for bullet in self.bullets.iter_mut() {
                        bullet.position.x = bullet.position.x + diff_pos_x;
                    }
                },
                Keycode::Z => {
                    self.shoot_main(ctx)?;
                },
                _ => {},
            }
        }

        let mut rng = rand::thread_rng();
        for star in self.stars.iter_mut() {
            star.pos.y += star.speed;
            if star.pos.y > self.window.height {
                star.pos.y = 0.;
                star.pos.x = Range::new(0., self.window.width).ind_sample(&mut rng);
                star.speed = Range::new(0.5 / 3., 4. / 3.).ind_sample(&mut rng);
            }
        }

        if self.music.stopped() {
            self.music.play()?;
        } else if self.music.paused() {
            self.music.resume();
        }

        //timer::yield_now();

        Ok(())
    }

    fn draw(&mut self, ctx: &mut Context) -> GameResult<()> {
        graphics::clear(ctx);

        self.draw_background(ctx)?;
        self.draw_stars(ctx)?;
        self.draw_player(ctx)?;

        let fps = format!("{:.*} bullets: {}", 1, timer::get_fps(ctx), self.bullets.len());
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

