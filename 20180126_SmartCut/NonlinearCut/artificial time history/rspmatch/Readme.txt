Ground Motion Spectral Matching Using Broyden Updating
Written by Armen Adekristi with support from Matthew Eatherton

ver1.0 Last update on: 26/4/2013

This routine performs spectral matching using the Corrected Tapered Cosine 
Wavelet coupled with Broyden Updating.

Example input is provided using the El Centro ground motion so that the file 
"SpectralMatching.m" can be opened and should run without any modifications.

The steps for using this routine are as follow:
    1. Open the File "SpectralMatching.m"
    2. Input the required information on the General Input, including:
       - the vector 'acc' which contains the acceleration time series (g)
       - the time increment of acceleration time series, named as 'dt'
       - the damping level of the SDOF oscillator, named as 'tetha'
       - the tolerance limit of average misfit (g) and maximum error, named 
         as 'avgtol'and 'errortol' respectively
       - the zero pad duration (sec) at ends of the acceleration time series,
         named as 'zeropad' 
    
    3. Input the required information on the Period Subset(s) and Target,
       including:
      - the period being matched (sec) in ascending order, named as 'Tall'  
      - the target response spectra (g) of the corresponding periods,named 
        as 'targetall'
      - the period ranges of each period subset, named as T1range, T2range
        and T3range
    
    4. Adjust the gain factors if desired
    
    5. Run the Analysis (F5)