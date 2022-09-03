use blackwinter2::{Game, GameResult};

use macroquad::prelude::*;

fn window_conf() -> Conf {
    Conf {
        window_title: "Black Winter II : Final Assault".to_owned(),
        window_width: 320,
        window_height: 413,
        ..Default::default()
    }
}

#[macroquad::main(window_conf)]
async fn main() -> GameResult {
    let conf = window_conf();
    let mut game = Game::new(conf.window_width as f32, conf.window_height as f32).await?;

    loop {
        if is_quit_requested() {
            break;
        }

        game.update().await?;
        game.draw().await?;

        next_frame().await
    }

    Ok(())
}
