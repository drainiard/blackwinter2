# Black Winter 2
A game originally made by [Wong Yat Seng][1], rewritten in Rust.

## Dev Notes
Since the original project was written in VB6, sizes like window width and
height were expressed in twips. To translate them in pixels I used this website:
 - [Convert twips to pixels][2]

---

**~ Finding a Graphics Library**
`2018-01-17`

**Attempt #1: Piston.rs**

Before starting the project I went to [arewegameyet](http://arewegameyet.com)
(a website that aggregates useful info and links to game-related Rust libraries
 and resources).

This first attempt didn't go very well, I tried running [an example][3] to see
how to instantiate a window and draw a sprite in it, but had troubles due to
[this issue][4].

**Attempt #2: ggez.rs**

I decided to give up and try ggez.rs which was also mentioned on a cool talk
(you can find the [video on youtube](5)) given by @lisapassing at RustFest
Zürich 2017.

ggez appears simpler to understand.

Sadly I'm [having issues here as well](6), apparently all Rust graphics libraries
depend on an openGL version that is not available on my OS right now.

---

[1]: http://www.a1vbcode.com/app-2021.asp
[2]: http://www.unitconversion.org/typography/twips-to-pixels-x-conversion.html
[3]: https://github.com/PistonDevelopers/piston-examples/blob/master/src/sprite.rs
[4]: https://github.com/PistonDevelopers/piston/issues/1202
[5]: https://www.youtube.com/watch?v=str_mex__0M
[6]: https://github.com/ggez/ggez/issues/194
