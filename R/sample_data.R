#' Sample Analytic Date for Creating Example Tables
#'
#' A dataset containing the soil data at one location with concentration of
#' volatile chemicals.
#'
#' @format A data frame with 235 rows and 20 variables:
#' \describe{
#'   \item{exposure_point}{area of exposure (ex. Project Area)}
#'   \item{medium}{sample medium type (ex. Soil)}
#'   \item{sample_material}{sample material type}
#'   \item{location_id}{unique key for sample location}
#'   \item{upper_depth}{upper depth range, if applicable}
#'   \item{lower_depth}{lower depth range, if applicable}
#'   \item{depth_units}{units for depth measurement, if applicable}
#'   \item{sample_date}{date sample was collected}
#'   \item{sample_id}{unique key for each sample}
#'   \item{sample_type}{additional sample information (ex. Field Dup)}
#'   \item{chem_class}{chemical class for analyte, can vary by project}
#'   \item{cas_rn}{CAS number for chemical with hyphens}
#'   \item{analyte}{short chemcial name}
#'   \item{analyte_name}{long chemical name}
#'   \item{concentration}{concentration of analyte in sample}
#'   \item{result_text}{text concentration of sample, typically concentration
#'   and flag}
#'   \item{units}{units for concentration}
#'   \item{meas_basis}{dry or wet weight, if applicable}
#'   \item{detected}{true if chemical was detected}
#'   \item{flags}{flags associated with sample concentration}
#'   \item{detect_limit}{detection limit for sample concentration}
#'   \item{detection_limit_units}{detection limit units}
#'   \item{table_num}{table number}
#'   \item{depth_bin}{upper_depth-lower_depth depth_units)}
#' }
