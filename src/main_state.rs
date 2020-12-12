use crate::bullet::{Bullet, Vector2D};
use crate::starfield::Starfield;
use ggez::audio;
use ggez::audio::SoundSource;
use ggez::event::{self, KeyCode, KeyMods};
use ggez::graphics;
use ggez::graphics::{Color, DrawParam, FilterMode};
use ggez::nalgebra::Point2;
use ggez::{timer, Context, GameResult};
use rand;
use rand::distributions::{IndependentSample, Range};
use std::collections::HashSet;
use std::path::Path;

const PLAYER_SPRITES_COUNT: u8 = 3;

pub struct RectSize {
    width: f32,
    height: f32,
}

pub struct MainState {
    font: graphics::Font,
    music: audio::Source,
    window: RectSize,
    pos: Point2<f32>,
    speed: f32,
    moving: HashSet<KeyCode>,
    stars: Vec<Starfield>,
    counter: u8,
    bullets: Vec<Bullet>,
    target_fps: u32,
}

impl MainState {
    pub fn new(ctx: &mut Context, (width, height): (f32, f32)) -> GameResult<Self> {
        ggez::filesystem::print_all(ctx);

        let font = graphics::Font::new(ctx, "/game_sans_serif_7.ttf")?;

        let mut rng = rand::thread_rng();
        let mut stars = Vec::with_capacity(100);
        for _ in 0..100 {
            let star = Starfield {
                pos: Point2::new(
                    Range::new(0., width).ind_sample(&mut rng),
                    Range::new(0., height).ind_sample(&mut rng),
                ),
                speed: Range::new(0.5 / 3., 4. / 3.).ind_sample(&mut rng),
                color: Range::new(0, 160).ind_sample(&mut rng),
            };
            stars.push(star);
        }

        let music = audio::Source::new(ctx, Path::new("/bgplay.ogg"))?;

        let target_fps = 60;

        let s = MainState {
            font: font,
            music: music,
            window: RectSize {
                width: width,
                height: height,
            },
            pos: Point2::new((width - 49.) / 2., (height - 38.) / 2.),
            speed: 3.0,
            moving: HashSet::new(),
            stars: stars,
            counter: 0,
            bullets: Vec::new(),
            target_fps,
        };
        Ok(s)
    }

    fn shoot_main(&mut self, ctx: &mut Context) -> GameResult {
        // TODO add logic to change behavior based on configured gun

        let image = graphics::Image::new(ctx, Path::new("/img0-1.png"))?;
        let bullet_image = graphics::Image::new(ctx, Path::new("/img6-1.png"))?;
        let adjust = Vector2D::new(
            (image.width() - bullet_image.width()) as f32 / 2.,
            bullet_image.height() as f32 * -1.,
        );

        self.bullets.push(Bullet::new(
            self.pos.x + adjust.x,
            self.pos.y + adjust.y,
            0.,
            self.speed * -6.,
        ));

        Ok(())
    }

    fn draw_stars(&mut self, ctx: &mut Context) -> GameResult {
        let mut mb = graphics::MeshBuilder::new();

        for star in self.stars.iter() {
            let color = Color::from_rgb(star.color, star.color, star.color);
            mb.line(
                &[
                    Point2::new(star.pos.x, star.pos.y),
                    Point2::new(star.pos.x, star.pos.y + 1.),
                ],
                1.,
                color,
            )?;
        }

        let mesh = mb.build(ctx)?;
        graphics::draw(ctx, &mesh, DrawParam::default())
    }

    fn draw_player(&mut self, ctx: &mut Context) -> GameResult {
        self.counter = (self.counter + 1) % (PLAYER_SPRITES_COUNT * 3);
        let counter: u8 = self.counter / 3;
        let image = graphics::Image::new(ctx, Path::new(&format!("/img0-{}.png", counter + 1)))?;
        let params: DrawParam = DrawParam::default().dest(Point2::new(self.pos.x, self.pos.y));
        graphics::draw(ctx, &image, params)?;

        let bullet_image = graphics::Image::new(ctx, Path::new(&format!("/img6-6.png")))?;

        let dead_bullets: Vec<usize> = self
            .bullets
            .iter()
            .enumerate()
            .filter(|&(_i, bullet): &(usize, &Bullet)| bullet.position.y < -413.)
            .map(|(i, _)| i)
            .rev()
            .collect();
        for i in dead_bullets {
            self.bullets.swap_remove(i);
        }

        let base_params: DrawParam = DrawParam::default();
        for bullet in self.bullets.iter() {
            let point = bullet.position.to_point2();
            graphics::draw(ctx, &bullet_image, base_params.dest(point))?;
        }

        Ok(())
    }
}

impl event::EventHandler for MainState {
    fn update(&mut self, ctx: &mut Context) -> GameResult {
        while timer::check_update_time(ctx, self.target_fps) {
            for bullet in self.bullets.iter_mut() {
                bullet.update(ctx)?;
            }

            let speed = self.speed;
            for keycode in self.moving.clone() {
                match keycode {
                    KeyCode::Up => {
                        let new_pos_y = (0. as f32).max(self.pos.y - speed);
                        let diff_pos_y = self.pos.y - new_pos_y;
                        self.pos.y = new_pos_y;
                        for bullet in self.bullets.iter_mut() {
                            bullet.position.y = bullet.position.y - diff_pos_y;
                        }
                    }
                    KeyCode::Down => {
                        let new_pos_y =
                            (self.window.height - 38. as f32).min(self.pos.y + self.speed);
                        let diff_pos_y = new_pos_y - self.pos.y;
                        self.pos.y = new_pos_y;
                        for bullet in self.bullets.iter_mut() {
                            bullet.position.y = bullet.position.y + diff_pos_y;
                        }
                    }
                    KeyCode::Left => {
                        let new_pos_x = (0. as f32).max(self.pos.x - self.speed);
                        let diff_pos_x = self.pos.x - new_pos_x;
                        self.pos.x = new_pos_x;
                        for bullet in self.bullets.iter_mut() {
                            bullet.position.x = bullet.position.x - diff_pos_x;
                        }
                    }
                    KeyCode::Right => {
                        let new_pos_x =
                            (self.window.width - 49. as f32).min(self.pos.x + self.speed);
                        let diff_pos_x = new_pos_x - self.pos.x;
                        self.pos.x = new_pos_x;
                        for bullet in self.bullets.iter_mut() {
                            bullet.position.x = bullet.position.x + diff_pos_x;
                        }
                    }
                    KeyCode::Z => {
                        self.shoot_main(ctx)?;
                    }
                    _ => {}
                }
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

    fn draw(&mut self, ctx: &mut Context) -> GameResult {
        graphics::clear(ctx, Color::from((0., 0., 0.)));

        self.draw_stars(ctx)?;
        self.draw_player(ctx)?;

        let fps = format!(
            "{:.*}\nbullets {}\nstars {}\nmusic {:?}",
            1,
            timer::fps(ctx),
            self.bullets.len(),
            self.stars.len(),
            self.music.elapsed()
        );
        let text_fragment = graphics::TextFragment::new(fps);
        let mut text = graphics::Text::new(text_fragment);
        text.set_font(self.font, graphics::Scale::uniform(10.));

        //graphics::draw(ctx, &text, DrawParam::default().dest(Point2::new(0., 0.)))?;

        graphics::queue_text(ctx, &text, [20., 20.], None);
        graphics::draw_queued_text(ctx, DrawParam::default(), None, FilterMode::Nearest)?;

        graphics::present(ctx)
    }

    fn key_down_event(
        &mut self,
        ctx: &mut Context,
        keycode: KeyCode,
        _keymods: KeyMods,
        _repeat: bool,
    ) {
        if let KeyCode::Escape = keycode {
            ggez::event::quit(ctx);
        }
        self.moving.insert(keycode);
    }

    fn key_up_event(&mut self, _ctx: &mut Context, keycode: KeyCode, _keymods: KeyMods) {
        self.moving.remove(&keycode);
    }
}
