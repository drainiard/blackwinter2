#[derive(Default, Debug)]
pub struct Starfield {
    pub pos_x: f32,
    pub pos_y: f32,
    pub speed: f32,
    pub color: f32,
}

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
    }
}
