#import GeneRead
import HLS_Algorithm
import HLS_Flowchart
import HLS_Text
import PuLP_Algorithm
import ATS_Algorithm
import ATS_Flowchart
import ATS_Text

Heuristic_Solution,delta_T=HLS_Text.HLS_Text()
Heuristic_Solution,delta_T=HLS_Flowchart.HLS_Flowchart()
Heuristic_Solution,delta_T=HLS_Algorithm.HLS_Algorithm()
Heuristic_Solution,delta_T=ATS_Text.ATS_Text()
Heuristic_Solution,delta_T=ATS_Flowchart.ATS_Flowchart()
Heuristic_Solution,delta_T=ATS_Algorithm.ATS_Algorithm()

# Cannot Compare with Heuristic and PuLP since PuLP uses VN but Heuristic has unlimited vehicles for each Type
#objec_val,Vehicle_Type_Maximum_Utilised_Capacity,Number_of_Vehicle_used_of_each_Type=PuLP_Algorithm.PuLP_Algorithm(max_seconds_allowed_for_calculation=delta_T)