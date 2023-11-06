# TCG (Test Case Generator)
Test Case Generator built for the USAF.  Some variable names and content have been changed to avoid sharing of classified information.  Example data and further information cannot be provided for the same reason.  This project was developed in a restricted computing environment - available packages were limited and there was no ability to download new packages.

## Project Scope
This project is designed to automate the generation of test cases for use in simulation testing.  The required inputs are: \
Test object list - .xlsx file which contains both coded identifiers and names for the systems being simulated  \
Simulation summary file(s) - .xlsx file(s) summarizing simulation(s) to be tested

With required inputs, the output will be a test case which contains 2 sheets per system to be simulated, the first sheet being a test case outlining the avaiable simulation modes with space to comment on the results of testing, the second sheet being the provided summary for quick reference.
