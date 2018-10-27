use std::ffi::OsString;
use std::fs::File;
use std::io::Write;

pub fn run(args: &[OsString]) -> Result<(), String> {
    let map = |err: std::io::Error| err.to_string();

    let mut outfile = File::create("C:\\ProgramData\\tasktest.txt").map_err(map)?;
    writeln!(outfile, "OK!").map_err(map)?;
    for (i, arg) in args.iter().enumerate() {
        writeln!(outfile, "{}: \"{}\"", i, arg.to_string_lossy()).map_err(map)?;
    }

    Ok(())
}
