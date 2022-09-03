use crate::bullet::Bullet;
use crate::error::GameError;
use crate::starfield::Starfield;
use macroquad::audio::{load_sound, play_sound, Sound};
use macroquad::{input, prelude::*};

pub struct RectSize {
    width: f32,
    height: f32,
}

pub struct Game {
    font: Font,
    music: Sound,
    current_music: Option<Sound>,
    window: RectSize,
    pos: Vec2,
    speed: f32,
    stars: Vec<Starfield>,
    counter: usize,
    bullets: Vec<Bullet>,
    bullet_image: Texture2D,
    player_images: Vec<Texture2D>,
}

pub type GameResult<T = ()> = Result<T, GameError>;

impl Game {
    pub async fn new(width: f32, height: f32) -> GameResult<Self> {
        let mut stars = Vec::with_capacity(100);
        for _ in 0..100 {
            let pos = Vec2::new(rand::gen_range(0., width), rand::gen_range(0., height));
            let speed = rand::gen_range(0.5 / 3., 4. / 3.);
            let color = rand::gen_range(0, 160);

            let star = Starfield { pos, speed, color };
            stars.push(star);
        }

        let font = load_ttf_font("./resources/game_sans_serif_7.ttf").await?;

        let music = load_sound("./resources/bgplay.ogg").await?;
        let current_music = None;

        let bullet_image = load_texture("./resources/img6-1.png").await?;
        let player_images = vec![
            load_texture("./resources/img0-1.png").await?,
            load_texture("./resources/img0-2.png").await?,
            load_texture("./resources/img0-3.png").await?,
        ];

        let s = Game {
            font,
            music,
            current_music,
            window: RectSize { width, height },
            pos: Vec2::new((width - 49.) / 2., (height - 38.) / 2.),
            speed: 3.0,
            stars,
            counter: 0,
            bullets: Vec::new(),
            bullet_image,
            player_images,
        };
        Ok(s)
    }

    async fn shoot_main(&mut self) -> GameResult {
        // TODO add logic to change behavior based on configured gun

        let image = self.player_images[0];
        let adjust = Vec2::new(
            (image.width() - self.bullet_image.width()) as f32 / 2.,
            self.bullet_image.height() as f32 * -1.,
        );

        self.bullets.push(Bullet::new(
            self.pos.x + adjust.x,
            self.pos.y + adjust.y,
            0.,
            self.speed * -6.,
        ));

        Ok(())
    }

    async fn draw_stars(&mut self) -> GameResult {
        for star in self.stars.iter() {
            let color = Color::from_rgba(star.color, star.color, star.color, 0xff);
            draw_line(
                star.pos.x,
                star.pos.y,
                star.pos.x,
                star.pos.y + 1.,
                1.,
                color,
            );
        }

        Ok(())
    }

    async fn draw_player(&mut self) -> GameResult {
        self.counter = (self.counter + 1) % (self.player_images.len() * 3);
        let counter = self.counter / 3;
        let image = self.player_images[counter];
        let bullet_image = self.bullet_image;

        draw_texture(image, self.pos.x, self.pos.y, WHITE);

        let dead_bullets: Vec<usize> = self
            .bullets
            .iter()
            .enumerate()
            // allow one bullet to overflow the window
            .filter(|&(_i, bullet): &(usize, &Bullet)| bullet.position.y < -bullet_image.height())
            .map(|(i, _)| i)
            .rev()
            .collect();
        for i in dead_bullets {
            self.bullets.swap_remove(i);
        }

        for bullet in self.bullets.iter() {
            let (x, y) = bullet.position.into();
            draw_texture(bullet_image, x, y, WHITE);
        }

        Ok(())
    }

    pub async fn update(&mut self) -> GameResult {
        if self.current_music.is_none() {
            self.current_music = Some(self.music);

            play_sound(
                self.music,
                macroquad::audio::PlaySoundParams {
                    looped: true,
                    ..Default::default()
                },
            );
        }

        for bullet in self.bullets.iter_mut() {
            bullet.update();
        }

        let speed = self.speed;

        if input::is_key_down(KeyCode::Up) {
            let new_pos_y = (0. as f32).max(self.pos.y - speed);
            let diff_pos_y = self.pos.y - new_pos_y;
            self.pos.y = new_pos_y;
            for bullet in self.bullets.iter_mut() {
                bullet.position.y = bullet.position.y - diff_pos_y;
            }
        }
        if input::is_key_down(KeyCode::Down) {
            let new_pos_y = (self.window.height - 38. as f32).min(self.pos.y + self.speed);
            let diff_pos_y = new_pos_y - self.pos.y;
            self.pos.y = new_pos_y;
            for bullet in self.bullets.iter_mut() {
                bullet.position.y = bullet.position.y + diff_pos_y;
            }
        }
        if input::is_key_down(KeyCode::Left) {
            let new_pos_x = (0. as f32).max(self.pos.x - self.speed);
            let diff_pos_x = self.pos.x - new_pos_x;
            self.pos.x = new_pos_x;
            for bullet in self.bullets.iter_mut() {
                bullet.position.x = bullet.position.x - diff_pos_x;
            }
        }
        if input::is_key_down(KeyCode::Right) {
            let new_pos_x = (self.window.width - 49. as f32).min(self.pos.x + self.speed);
            let diff_pos_x = new_pos_x - self.pos.x;
            self.pos.x = new_pos_x;
            for bullet in self.bullets.iter_mut() {
                bullet.position.x = bullet.position.x + diff_pos_x;
            }
        }
        if input::is_key_down(KeyCode::Z) {
            self.shoot_main().await?;
        }

        for star in self.stars.iter_mut() {
            star.pos.y += star.speed;
            if star.pos.y > self.window.height {
                star.pos.y = 0.;
                star.pos.x = rand::gen_range(0., self.window.width);
                star.speed = rand::gen_range(0.5 / 3., 4. / 3.);
            }
        }

        Ok(())
    }

    pub async fn draw(&mut self) -> GameResult {
        clear_background(BLACK);

        self.draw_stars().await?;
        self.draw_player().await?;

        draw_rectangle(8., 8., 70., 40., Color::from_rgba(0, 0, 255, 100));

        let params = TextParams {
            color: WHITE,
            font: self.font,
            font_size: 10,
            ..Default::default()
        };
        draw_text_ex(&format!("{:.2}", 1. / get_frame_time()), 20., 20., params);
        draw_text_ex(&format!("bullets {}", self.bullets.len()), 20., 30., params);
        draw_text_ex(&format!("stars {}", self.stars.len()), 20., 40., params);

        Ok(())
    }
}
