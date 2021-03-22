# MaximumLikelihoodGammaDist
A basic implementation for the maximum likelihood estimators of a gamma distribution's parameters.


The implementation is based on Thomas P. Minka's algorithm for the maximum likelihood estimation of a gamma distribution's parameters and Jose Bernardo's algorithms to calculate the values of the digamma and the trigamma functions for positive input variables and John Burkardt's implementations respectively.
See the following references:

[Minka - Estimating a Gamma Distribution](https://tminka.github.io/papers/minka-gamma.pdf)

[Burkardt - The Digamma or Psi function](https://people.math.sc.edu/Burkardt/py_src/asa103/asa103.html)

[Burkardt - The Trigamma function](https://people.math.sc.edu/Burkardt/f_src/asa121/asa121.html)


The shape parameter **α** can be estimated by using the function *GammaMLAlpha*. The input is a one dimensional range in a worksheet containing only positive numbers which are suspected to follow a gamma distribution. Optionally a threshold can be given which stops the itrative calculation of α when undercut.


The scale parameter **β** can be estimated by using the function *GammaMLBeta*, which takes the mean of a sample and a shape parameter α as input. While it is perfectly possible to use arbitrary values, it is recommended to use the mean of the same sample the shape parameter is estimated from. The mean must be positive.
