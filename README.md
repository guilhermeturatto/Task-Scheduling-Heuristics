# Task Scheduling Heuristics

Scheduling problems arise when it is necessary to define the sequence of task execution in order to minimize costs or losses resulting from delays or advancements in the end of execution in relation to the established deadline.

<p align="center">
<img src="https://imgur.com/HYrgDLX.png" alt="drawing" width="400"/>
</p>

This approach can be used in various areas, such as Project Management, Production Planning, Machinery Allocation and Logistics. The scheduling problem can be modeled with 4 parameters for each task:

* <code>d<sub>j</sub></code> Task j's deadline
* <code>p<sub>j</sub></code> Task j's processing time
* <code>h<sub>j</sub></code> Task j's penalty for advancement
* <code>w<sub>j</sub></code> Task j's penalty for delay

<p align="center">
<img src="https://imgur.com/1wbrJo8.png" alt="drawing" width="650"/>
</p>

The solution consists of determining the execution sequence that generates the lowest possible penalty, considering that all tasks will be executed sequentially without parallelism. These problems are classified as NP-hard, therefore, ILP (Integer Linear Programming) models only apply to problems with few tasks. For larger dimension problems, heuristic methods are often used.

In this project, the following heuristics were implemented:

>* **SPT (Shortest Processing Time)**: Process tasks with the shortest processing time first <code>p<sub>1</sub> ≤ p<sub>2</sub> ≤ ... ≤ p<sub>n</sub></code>.
>
>* **LPT (Longest Processing Time)**: Process tasks with the longest processing time first <code>p<sub>1</sub> ≥ p<sub>2</sub> ≥ ... ≥ p<sub>n</sub></code>.
>
>* **WSPT (Weighted Shortest Processing Time)**: Process tasks with the smallest ratio between the fine for delay and the processing time first <code>(w<sub>1</sub>/p<sub>1</sub>) ≥ (w<sub>2</sub>/p<sub>2</sub>) ≥ ... ≥ (w<sub>n</sub>/p<sub>n</sub>)</code>.
>
>* **WLPT (Weighted Longest Processing Time)**: Process tasks with the smallest ratio between the fine for advancement and the processing time first <code>(h<sub>1</sub>/p<sub>1</sub>) ≤ (h<sub>2</sub>/p<sub>2</sub>) ≤ ... ≤ (h<sub>n</sub>/p<sub>n</sub>)</code>.
>
>* **EDD (Erliest Due Date)**: Process tasks with the earliest deadline first <code>d<sub>1</sub> ≤ d<sub>2</sub> ≤ ... ≤ d<sub>n</sub></code>.
>
>* **MST (Minimum Slack Time)**: Process tasks with the smallest difference between the deadline and the processing time first \
><code>(d<sub>1</sub>–p<sub>1</sub>) ≤ (d<sub>2</sub>–p<sub>2</sub>) ≤ ... ≤ (d<sub>n</sub>–p<sub>n</sub>)</code>.

The performance of each heuristic may vary according to the characteristics of the set of tasks. 


## Usage

### Windows

In windows, the program can be run in two ways:

- Run the self-contained ```SchedulingHeuristics.exe``` to open the interface (may take a few seconds).

or

- Open a command prompt CMD in the folder containing "SchedulingHeuristics.py" and run:
 ```
 pip install -r requirements.txt
 ```
 ```
 py SchedulingHeuristics.py
 ```
 > OBS: In this case, [Python](https://www.python.org/) and [pip](https://pypi.org/project/pip/) are required. [Virtual environment](https://docs.python.org/3/library/venv.html) is recommended.

### Linux

- Open a terminal in the folder containing "SchedulingHeuristics.py" and run:
 ```
 pip install -r requirements.txt
 ```
 ```
 py SchedulingHeuristics.py
 ```
 > OBS: In this case, [Python](https://www.python.org/) and [pip](https://pypi.org/project/pip/) are required. [Virtual environment](https://docs.python.org/3/library/venv.html) is recommended.

---

### OBS:

* The spreadsheet containing the task data must be provided in the same folder that the program;
* The spreadsheet must be closed when click in "Start";
* After "Calculations Completed" message, the spreadsheet can be opened and the results will be updated;
* The number and name of tasks can be modified;
* The fine and time units must match. Ex: Deadline in days, fine per day. Deadline in hours, fine per hour. 

Expected result:
<p align="center">
<img src="https://imgur.com/PNatZrA.png" alt="drawing" width="900"/>
</p>

## Support
If you want some help with this work contact me: guilherme.turatto@gmail.com

