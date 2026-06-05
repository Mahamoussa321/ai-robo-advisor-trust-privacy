# Global soil pollution by toxic metals threatens agriculture and human health

[https://doi.org/10.5061/dryad.83bk3jb2z](https://doi.org/10.5061/dryad.83bk3jb2z)

## Description of the data and file structure

#### File: Attachment\_1\_Soil\_pollution\_data\_sources

**Description:** This table compiles literature selected for the global soil pollution dataset. It presents key details, including the first author's last name, publication year, region, study area coordinates, and the title of the literature in separate columns. Detailed reference information for these publications is located at the end of the table.

#### File: Attachment\_2\_Source\_of\_co-variate\_and\_variable\_retention\_for\_modeling

**Description:** This table provides basic information on variables used in modeling. The first column lists all the variables used in the model. Columns 2-8 show variables remaining after feature selection under Human health and ecological thresholds for each toxic metal. Columns 9-15 display variables remaining after feature selection under Agricultural thresholds for each toxic metal. If the variable is selected, it will be marked with a Y, otherwise it will be blank. Columns 16-22 present detailed information about the datasets of covariates, including the spatial resolution, temporal coverage, unit of measurement, data format, database version, data source and download link for each variable. The empty cells and symbol “\” both indicate that the corresponding dataset does not contain the specified information, i.e. missing data.

#### File: Attachment\_3\_Feature\_importance\_and\_correlation\_with\_toxic\_metals

**Description:** This table provides the importance of variables calculated using Shapley Additive Explanations (SHAP) and Mean Decrease in Node Impurity (MDI), along with the variables' correlation with toxic metals. It displays the importance and correlation of variables remaining in models after feature selection. The empty cell indicates that the corresponding variable has not been selected in the feature selection. Columns 1-2 list variable names and abbreviations used in modeling. Columns 3-9 and Columns 10-16 respectively show the average absolute value of SHAP value in models for toxic metals under Human Health and Ecological Thresholds (HHET) and Agricultural Thresholds (AT). Higher values indicate greater variable importance in the model. Columns 17 -23 and Columns 24-30 display the Pearson correlation coefficients between selected variables and concentrations of toxic metals in all the land use types and agricultural lands, respectively. Columns 31-37 and Columns 38-44 reveal the importance of variables calculated by MDI under HHET and AT. Larger values indicate greater variable importance.

#### File: Attachment\_4\_Global\_dataset\_of\_predicted\_toxic\_metals\_exceedance\_under\_HHET

**Description:** This table presents the probability of exceedance for different toxic metals under Human Health and Ecological Thresholds for each grid. Columns 1-2 provide the coordinates of the grid, and probabilities of exceedance for toxic metals are listed in the following columns.

#### File: Attachment\_5\_Global\_dataset\_of\_predicted\_toxic\_metals\_exceedance\_under\_AT

**Description:** This table presents the probability of exceedance for different toxic metals under Agricultural Thresholds for each grid. Columns 1-2 provide the coordinates of the grid, and probabilities of exceedance for toxic metals are listed in the following columns.

#### File: Attachment\_6\_Distribution\_of\_sample\_location\_for\_As

**Description:** This figure presents the distribution of sample locations for As. The color of points represents the number of samples in each location. Points data (e.g. data from LUCAS) were aggregated according to administration devotions.

#### File: Attachment\_7\_Distribution\_of\_sample\_location\_for\_Cd

**Description:** This figure presents the distribution of sample locations for Cd. The color of points represents the number of samples in each location. Points data (e.g. data from LUCAS) were aggregated according to administration devotions.

#### File: Attachment\_8\_Distribution\_of\_sample\_location\_for\_Co

**Description:** This figure presents the distribution of sample locations for Co. The color of points represents the number of samples in each location. Points data (e.g. data from LUCAS) were aggregated according to administration devotions.

#### File: Attachment\_9\_Distribution\_of\_sample\_location\_for\_Cr

**Description:** This figure presents the distribution of sample locations for Cr. The color of points represents the number of samples in each location. Points data (e.g. data from LUCAS) were aggregated according to administration devotions.

#### File: Attachment\_10\_Distribution\_of\_sample\_location\_for\_Cu

**Description:** This figure presents the distribution of sample locations for Cu. The color of points represents the number of samples in each location. Points data (e.g. data from LUCAS) were aggregated according to administration devotions.

#### File: Attachment\_11\_Distribution\_of\_sample\_location\_for\_Ni

**Description:** This figure presents the distribution of sample locations for Ni. The color of points represents the number of samples in each location. Points data (e.g. data from LUCAS) were aggregated according to administration devotions.

#### File: Attachment\_12\_Distribution\_of\_sample\_location\_for\_Pb

**Description:** This figure presents the distribution of sample locations for Pb. The color of points represents the number of samples in each location. Points data (e.g. data from LUCAS) were aggregated according to administration devotions.

#### File: Attachment\_13\_Code\_for\_models\_development\_and\_data\_analysis

**Description:** This file contains the code used in this study for model development and data analysis. The code is primarily written in Python, and the utilized packages are listed.
