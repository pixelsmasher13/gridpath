//! Debug-timing helper for the `IC_DEBUG=1` phase logs. `std::time::Instant`
//! aborts on `wasm32-unknown-unknown` (no clock), so the timer is inert there.
pub struct DebugTimer {
    #[cfg(not(target_arch = "wasm32"))]
    start: std::time::Instant,
}
impl DebugTimer {
    pub fn start() -> Self {
        DebugTimer {
            #[cfg(not(target_arch = "wasm32"))]
            start: std::time::Instant::now(),
        }
    }
    /// Elapsed milliseconds (always 0 on wasm32).
    pub fn ms(&self) -> u128 {
        #[cfg(not(target_arch = "wasm32"))]
        { self.start.elapsed().as_millis() }
        #[cfg(target_arch = "wasm32")]
        { 0 }
    }
}
