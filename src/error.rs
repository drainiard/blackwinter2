use macroquad::prelude::*;

#[derive(Debug)]
pub struct GameError(String);

impl From<FontError> for GameError {
    fn from(e: FontError) -> Self {
        Self(e.to_string())
    }
}

impl From<FileError> for GameError {
    fn from(e: FileError) -> Self {
        Self(e.to_string())
    }
}
