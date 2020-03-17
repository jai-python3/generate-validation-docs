#!/usr/bin/env bash
DIRNAME=$(dirname "$0")
source $DIRNAME/venv/bin/activate
python $DIRNAME/generate_validation_docs.py "$@"
