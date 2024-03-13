# CS 130

## Setting up a VENV

You can optionally set up a venv to run this code to prevent it from
interfering with other python packages on your system.

Create one in the directory `venv_dir` using...

```
python3 -m venv venv_dir
```

Then activate it using (assuming you're on macos/linux)...

```
. ./venv_dir/bin/activate
```

## Getting the Required Packages

Install the required packages using...

```
pip3 install -r requirements.txt
```

## Running Tests

Run all the tests using...

```
python3 -m unittest
```

You can generate profiles using `cProfile` for diagnosing performance...

```
python3 -m cProfile -o out.profile test/test_stress.py
```

