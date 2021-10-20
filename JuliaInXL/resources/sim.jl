using Distributions
using JuliaWebAPI
using Logging
Logging.configure(level=DEBUG)

f(u) = exp(-u^2/2)/√(2pi)

const u1=Uniform(-5.0,+5.0)
const u2=Uniform(0.0, 0.5)

function simulate_pt()
    x=rand(u1)
    y=rand(u2)
    y<f(x)
end

function simulate(num::Int)
     hits=0
     for (i in 1:num)
        hits = hits + simulate_pt()
    end
    return hits/num*(0.5*10)
end

function simulateTime(num)
    tic()
    r = simulate(num)
    t = toc()
    return (r, t)
end

simulate(num::String) = simulate(parseint(num))

#process([(simulate,false), (simulateTime,false)], "tcp://127.0.0.1:9999"; bind=true)