Compare-Seismic-Analysis-Methods
================================

These short scripts are being prepared for use in the preparation of a forthcoming white paper comparing the seismic analysis methods dictated by USNRC Regulatory Guide 1.92.

The analysis includes performing a response spectra (RS) and time history (TH) seismic analysis on a SAP2000 model created specifically for this comparison. The forces for a single column, modeled using prismatic frame elements, are extracted by the script and combined using the methods indicated in the U.S. Nuclear Regulatory Commisions Regulatory Guide 1.92, Combining Modal Response And Spatial Components In Seismic Response Analysis. For the RS analysis, these methods include the Square Root Sum of Squares method (SRSS) and the 100-40-40 rule. For the TH analysis, these methods include the 100-40-40 rule and Algebraic Sum (AS).

This script can extract and combine the forces resulting from either type of analysis (RS or TH) using any of the above mentioned methods, SRSS, 100-40-40, or AS.

After combining the forces, the script will then prepare the required number of pcaColumn Text Input (CTI) files in order to proceed with column design using pcaColumn v4.10. pcaColumn is used in order to arrive at a final design scheme that can be used for comparison in the forthcoming paper. Please note that this script is in no way intended to be used in any sort of analysis or design related to nuclear structures in the United States.
