# Call Julia functions from Microsoft Excel

### Installation

The JuliaInXL system is composed of two parts, a Julia package and an Excel Plugin. The Excel plugin itself consists of an Microsoft Excel add-in.
The Excel add-in is included in the distribution of the Julia package, so you only need to follow these steps:

0. If you don't have julia installed on your Windows system, follow [these](https://github.com/JuliaLang/juliaup#windows) instructions to get it. We only support the Windows distribution (no WSL yet, but stay tuned).
1. `import Pkg; Pkg.add("JuliaInXL")` (or `using JuliaInXL` and reply `yes` to the prompt to install)
2. If the installer doesn't popup, we need to trigger the package build step again (usually happens automatically, but may fail if you have tried multiple times unsuccessfully): `import Pkg; Pkg.build()`.
3. You should be seeing an installer by now. Click yes to all.
4. Close the julia window, open Excel and go to the JULIA tab
5. Click Launch Local Julia
6. Use `jleval` to run julia code


#### Useful tips

1. Typing something in `=jleval("JULIA PROGRAM")` is like typing in the repl. You can try doing `=jleval("x = 1")` and then investigate x in the REPL
2. Excel only triggers Julia execution once - if you have `B2` `=jleval("rand()") + A1` and change the `A1` cell, B2 will not update. 
3. `jlcall` only supports floats and strings. Passing integers will fail. It's mostly designed to be fast. It's recommended to stick with `jleval` for maximum flexibility.
4. Try to separate the logic between Excel and Julia. While it can be very satisfying to use excel as a REPL, it's better to keep a reproducible analysis in a `.jl` file and invoke it (possibly with special parameters) from Excel. That way you get the best of both worlds, without making entangled [spaghetti](https://en.wikipedia.org/wiki/Spaghetti_code). JuliaInXL shines most in the `communication` between Julia and Excel; if you find yourself using it as an IDE, you may want to try the [Julia Extension a try](https://code.visualstudio.com/docs/languages/julia) or [Pluto.jl](https://github.com/fonsp/Pluto.jl)!

### Advanced Usage Workflow

#### Getting Started
The primary supported workflow is for interactive development of Julia programs alongside Excel. Once the packages are installed, start Excel, and a Julia REPL. On the Julia repl, type `using JuliaInXL` to make the environment ready for calling via Excel. After that, define your functions as normal on the REPL. To try an an example, use a demo file included with the package: `include("C:\\Users\\<username>\\AppData\\Local\\JuliaPro-W.X.Y.Z\\Julia-W.X.Y\\share\\julia\\site\\vW.X\\JuliaInXL\\test\\sim.jl")` where `<username>` is your current Windows username and W.X.Y.Z are integers representing the current release of Julia Professional. (When using theses examples, you will need the `Distributions.jl` package installed. Please run `Pkg.add("Distributions")` on the Julia REPL if you do not have this already installed.)

Once you have the functions that you need to call from Excel, expose them via the `process_async` call. This will start the server listening to requests from Excel in the background, and return a connection object that can be used for exposing more functions later. The primary arguments to this function is an array of functions to be exposed, and the endpoint on which to listen to messages. The `bind=true` parameter tells the function to start listening at the endpoint and wait for client connections. 

```julia
conn = process_async([simulate, simulateTime], "tcp://127.0.0.1:9999", bind=true)
```

There is a corresponding `process` function with exactly the same API which listens synchronously, and thus doesn't return. This function should be used if the process is started from a command line, rather than via an iteractive REPL. 

#### Calling functions from Excel

Once the server is started, julia functions can be called from Excel using the `jlcall` worksheet function. The first argument to jlcall is a string, which is the name of the function to be called. Subsequent arguments to the `jlcall` function are passed as parameters to the Julia function being called. These can be constant literals, or cell references. Arrays can be passed via cell ranges. 

If the Julia function returns an array (1d or 2d), then use `jlcall` as an Excel Array function by selecting a range before entering the function, and pressing `Shift-Ctrl-Enter` to finish.

Functions exposed to Excel should take floats or strings, or their arrays as arguments. In general, it is a good idea to keep the function arguments as loosely typed as possible. Therefore functions should return integers, floats, or strings; or their arrays. However, arrays of dimensions greater than two are not supported. 

Note that [Excel stores all numbers as 64 bit IEEE floats](https://support.microsoft.com/en-us/kb/78113). Therefore, be aware of the possibility of truncation if returning large, or high precision, numbers. 

Dates are passed in from excel as floating point numbers in its internal encoding (fractional days since 1/1/1900 or 1/1/1904). Thus, they are recieved in Julia functions as floats. They can be converted to Julia DateTime values using the `xldate` function. 

#### Making changes

Remember that we started the Julia server in async mode. This means that the REPL is available for interactive use when working with Excel. Creating new definitions of functions that are already exposed will replace them, and the new versions will get called when the sheet is next recalculated. 

New functions can be added to the listener interface using the connection object saved from the original `process_async` call, via the `register` function. The arguments to `register` are the connection object returned from `process_async`, and the function name to expose. 

```julia
register(conn, simulateArray)
```
#### Controlling Julia from within Excel
![Ribbon](https://raw.githubusercontent.com/JuliaComputing/JuliaInXL/master/docs/addin-ribbon.png?token=AAXIJjVyMx7f5eYINZh9p0OAMleG68Luks5WmFXAwA%3D%3D)

----
_Microsoft and Excel are registered trademarks of Microsoft Corporation_
