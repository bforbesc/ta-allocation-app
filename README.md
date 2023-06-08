# 🧩 Teaching Assistants Allocation
The purpose of this application is to assign Teaching Assistants (TAs) to the faculty courses, taking into account various factors such as their contract percentage, individual preferences, and the specific needs of each course. 
By analyzing this information, the app streamlines the allocation process, ensuring a fair and optimized distribution of TAs across different courses. 

The app requires the following pieces of information, which are accepted as Excel files:
1. **Faculty courses**: list of courses for (bachelor's and master's) with respective number of classes and expected number of students
1. **BS course weights**: Bachelor's courses weights (conversion from needs to TAs contract percentage)
1. **TAs capacity**: TAs current contract percentage (from previous semester)
1. **TAs preferences**: TAs course and contract preferences

The app generates multiple outputs, including intermediate reports of data inconsistencies, culminating in the following main outputs:
1. **Course needs**: workload required for all ```BSC```, ```MST``` and ```ME``` courses
1. **Cleaned TAs preferences**: TAs ranked course preferences cleaned
1. **Automatic allocation results**: Results for automatic allocations for first preferences for both bachelor's and masters' courses

You can find the app here: [bforbesc-clustering-web-app-ml-web-app-ee5tk5.streamlit.app](https://bforbesc-ta-allocation-app-ta-allocation-app-m2v0xg.streamlit.app/)
