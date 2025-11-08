# Computation Offloading Based on DQN

This project implements a Deep Q-Network (DQN) reinforcement learning approach for computation offloading decisions in Mobile Edge Computing (MEC) environments. It compares DQN against traditional Q-learning.

## Requirements

- Python 3.7+
- PyTorch
- NumPy
- Matplotlib
- Pandas

## Installation

1. Install the required packages:

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install torch numpy matplotlib pandas
```

## Running the Project

Simply run the main training script:

```bash
python run_dqn.py
```

This will:
- Train a DQN agent and a Q-learning agent for 30,000 episodes
- Compare their performance on computation offloading tasks
- Display plots showing:
  - Average delay 
  - Average reward
  - Training progress over time

## Project Structure

- `run_dqn.py` - Main training script
- `dqn.py` - DQN neural network implementation
- `env.py` - Environment simulation (MEC servers, tasks, channels)
- `Q_table.py` - Traditional Q-learning implementation
- `func.py` - Action transformation utilities
- `constants.py` - Configuration parameters
- `task.py` - Task class definition

## How It Works

The system simulates a mobile device that must decide whether to:
1. Process tasks locally
2. Offload tasks to one of 3 MEC servers

The agents learn to optimize task completion time by considering:
- Task data size
- MEC server computational capacity
- Wireless channel conditions
- Network transmission delays

