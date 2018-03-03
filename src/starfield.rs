use ggez::graphics::Point2;

#[derive(Debug)]
pub struct Starfield {
    pub pos: Point2,
    pub speed: f32,
    pub color: u8,
}
