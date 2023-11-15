use std::path::PathBuf;

fn main() {
    let libxl_base = "D:/Projects/Components/LibXL4200";
    let header_file=format!("{libxl_base}/include_c/libxl.h");

    // Tell cargo to look for shared libraries in the specified directory
    println!("cargo:rustc-link-search={libxl_base}/lib64");

    // Tell cargo to tell rustc to link the libxl library.
    println!("cargo:rustc-link-lib=libxl");

    // Tell cargo to invalidate the built crate whenever the source file changes
    println!("cargo:rerun-if-changed={header_file}");

    // The bindgen::Builder is the main entry point
    // to bindgen, and lets you build up options for
    // the resulting bindings.
    let bindings = bindgen::Builder::default()
        // The input header we would like to generate
        // bindings for.
        .header(header_file)
        // Tell cargo to invalidate the built crate whenever any of the
        // included header files changed.
        .parse_callbacks(Box::new(bindgen::CargoCallbacks::new()))
        // Finish the builder and generate the bindings.
        .generate()
        // Unwrap the Result and panic on failure.
        .expect("Unable to generate bindings");

    //println!("cargo:rustc-env=OUT_DIR=./src");
    // Write the bindings to the $OUT_DIR/bindings.rs file.
    let out_path = PathBuf::from("./src");
    bindings
        .write_to_file(out_path.join("bindings.rs"))
        .expect("Couldn't write bindings!");
}
