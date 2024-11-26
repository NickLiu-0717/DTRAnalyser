# DTRAnalyser
Personal project to analyze daily training report(DTR) for imporving training drug detection dog

## Outline
1. What is a **Daily Training Report(DTR)**?
2. What potential issues in the DTR can we address to improve dog's performance?
3. What functionalities do we expect from this project?

### 1. What is a daily training report?
  - A daily training report is a document where a handler records their drug detection dog's perfromance everyday.
  - Without an odor source, a normal operation includes basic details such as duty time, duration, layout of the duty location, and dog's performance.
  - With an odor source, the training session includes basic details mentioned above, as well as the type of odor source, the amount of odor source, diffusion time, the method of setting up the odor source, and the purpose of the training session.
  - We grade each operation as **G**, **F**, **P**, **DNF**, or **FR**, which stand for **Good**, **Fair**, **Poor**, **Do Not Found**, and **False Response**.
  - If a dog finds a target ordor scent, it is considered a reward, we calculate the **Reward Ratio** to indicate how many operations it takes for the dog to find a reward.
  - In addition to real drugs odor source, we also use **Scented ordor sources**, which are materials of various kinds infused with the scent of drugs.
  - We record whether a dog shows a **Change In Behavior(CIB)** during an operation, specifying the potential cause of the **CIB**.
  - More details such as how many operations per day, where the operation takes, what's the reward ratio for everything duty location.

### 2. What potential issues in the DTR can we address to improve dog's performance?
  - **Time**: Both duty time and duration of the operation affect dog's performance, a fixed duty time or duration might lead to a pattern where dogs can anticipate whether they will receive a reward or not, which futher determines the effort they put into the operation.
  - **Odor Source**: Odor source is the key factor to train dogs. Variation in **the type of odor**, **amount**, **diffusion time**, **the method of settin up the source**, and **the freqency of odor usage** all play an important role.
  - **Reward Ratio**: Reward is the primary motivation for actively sniffing cargo. Too many consecutive operations without reward might reduce a dog's intent for reward. An optimal reward ration should be around 2.5 to 3.5. A high reward ratio might result in a lack of training and a reduction in motivation, while a low reward ration might make a dog overly excited from excessive anticipation, leading to poor sniffing quality.
  - **Training Cargo**: Dogs need to learn to sniff every types of cargo. The reward frequency for dogs in specific cargo types is proportional to their willingness to sniff them.
  - **Training Purpose**: Each operation with an odor source should have a specific training purpose to teach the dog. Providing sufficient training for each type of purpose is crucial to maintaining consistent improment in the dogs.

### 3. What functionalities do we expect from this project?
  - Analyze that if the duty time or duty duration is too consistent, recommandation for changes.
  - Analyze the usage frequency of each type of odor source, indicate how many times an odor source has been used over a specific period(Draw a statistical chart). Track changes in the usage of different amounts of each odor source(draw a trend graph).
  - Calculate the reward ratio over a specific period and analyze whether it correlates with the dog's performance.
  - Analyze how many types of cargo have been trained and determine whether the training frequency for each type is optimal.
  - List the training purpose for a given period, along with the corresponding date and the number of times each purpose was conducted.
