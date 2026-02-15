use chrono::{DateTime, NaiveTime};

pub struct TimePeriod {
    start:          NaiveTime,
    end:            NaiveTime,
    wraps_midnight: bool,
    name:           &'static str,
}

impl TimePeriod {
    pub const fn new(start: NaiveTime, end: NaiveTime, wraps_midnight: bool, name: &'static str) -> Self {
        Self {
            start,
            end,
            wraps_midnight,
            name,
        }
    }

    pub fn contains(&self, time: &NaiveTime) -> bool {
        if self.wraps_midnight {
            time >= &self.start || time < &self.end
        } else {
            time >= &self.start && time < &self.end
        }
    }
}

pub const fn hms(h: u32, m: u32) -> chrono::NaiveTime {
    chrono::NaiveTime::from_hms_opt(h, m, 0).unwrap()
}

// TODO: [trait] : We should do this via Trait + blank structs
// Which will be better as adding a 'new' period would be easier via ( new struct + impl Trait )
pub const PERIODS: [TimePeriod; 9] = [
    TimePeriod::new(hms(5, 0), hms(8, 45), false, "Before School"),
    TimePeriod::new(hms(8, 45), hms(8, 55), false, "Form"),
    TimePeriod::new(hms(8, 55), hms(10, 5), false, "Period 1"),
    TimePeriod::new(hms(10, 5), hms(11, 15), false, "Period 2"),
    TimePeriod::new(hms(11, 15), hms(11, 55), false, "Morning Tea"),
    TimePeriod::new(hms(11, 55), hms(13, 5), false, "Period 3"),
    TimePeriod::new(hms(13, 5), hms(13, 45), false, "Second Lunch"),
    TimePeriod::new(hms(13, 45), hms(14, 55), false, "Period 4"),
    TimePeriod::new(hms(14, 55), hms(5, 0), true, "After Hours"),
];

pub(crate) fn get_current_period(
    now: &DateTime<chrono::Local>,
    periods: &[TimePeriod],
) -> Result<String, String> {
    let current = now.time();

    for period in periods.iter() {
        if period.contains(&current) {
            return Ok(period.name.to_string());
        }
    }

    Err(format!("Current time {} does not fall into any defined period", current))
}
