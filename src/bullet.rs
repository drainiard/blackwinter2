use macroquad::prelude::*;

#[derive(Debug)]
pub struct Bullet {
    pub position: Vec2,
    pub direction: Vec2,
}

impl Bullet {
    pub fn new(pos_x: f32, pos_y: f32, dir_x: f32, dir_y: f32) -> Self {
        Self {
            position: Vec2::new(pos_x, pos_y),
            direction: Vec2::new(dir_x, dir_y),
        }
    }

    pub fn update(&mut self) {
        self.position.x = self.position.x + self.direction.x;
        self.position.y = self.position.y + self.direction.y;
    }
}
