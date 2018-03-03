extern crate rand;
extern crate ggez;

mod main_state;
mod bullet;
mod starfield;

pub use main_state::MainState;

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
    }
}
