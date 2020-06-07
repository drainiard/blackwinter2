use ggez::{nalgebra::Point2, Context, GameResult};

#[derive(Default, Debug)]
pub struct Vector2D {
    pub x: f32,
    pub y: f32,
}

impl Vector2D {
    pub fn new(x: f32, y: f32) -> Self {
        Self { x: x, y: y }
    }

    pub fn add(&self, direction: &Vector2D) -> Self {
        Self::new(self.x + direction.x, self.y + direction.y)
    }

    pub fn to_point2(&self) -> Point2<f32> {
        Point2::new(self.x, self.y)
    }
}

#[derive(Debug)]
pub struct Bullet {
    pub position: Vector2D,
    pub direction: Vector2D,
}

impl Bullet {
    pub fn new(pos_x: f32, pos_y: f32, dir_x: f32, dir_y: f32) -> Self {
        Self {
            position: Vector2D::new(pos_x, pos_y),
            direction: Vector2D::new(dir_x, dir_y),
        }
    }

    pub fn update(&mut self, _ctx: &mut Context) -> GameResult {
        self.position.x = self.position.x + self.direction.x;
        self.position.y = self.position.y + self.direction.y;

        Ok(())
    }
}
