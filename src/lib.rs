mod bullet;
mod error;
mod main_state;
mod starfield;

pub use main_state::Game;
pub use main_state::GameResult;

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
    }
}
