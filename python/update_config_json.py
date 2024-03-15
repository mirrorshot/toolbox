#!/usr/bin/python3

import json
import os
import socket
from argparse import ArgumentParser


def compose_argument_parser(argument_parser: ArgumentParser = None) -> ArgumentParser:
    argparser: ArgumentParser = argument_parser or ArgumentParser(
        prog=os.path.basename(__file__),
        description="""
        Script used to perform an in-place update of a config file for the console IDE plugin. 
        """,
    )
    argparser.add_argument(
        "config_path",
        type=str
    )
    argparser.add_argument(
        "-o", "--ollama_ip",
        type=str,
        required=False,
        default="local",
        help="If omitted or given value \"local\" makes the script lookup the IP of the machine running the script, useful in WSL."
    )
    argparser.add_argument(
        "--ollama_port",
        type=int,
        required=False,
        default="11434"
    )
    return argparser


def get_ip() -> str:
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.settimeout(0)
    try:
        s.connect(('8.8.8.8', 1))
        IP = s.getsockname()[0]
    except Exception as e:
        raise Exception("failed extracting local IP") from e
    finally:
        s.close()
    return IP


def update_config(config_path: str, ollama_ip: str, ollama_port: str) -> None:
    try:
        with open(config_path, "r") as config_file:
            console_config = json.load(config_file)

        for model in console_config["models"]:
            if model["provider"] == "ollama":
                old_api_base = f"http://{ollama_ip}:{ollama_port}"
                model["apiBase"] = old_api_base
                print(f"old url {old_api_base} -> updated ollama model: {json.dumps(model)}")

        with open(config_path, "w") as config_file:
            json.dump(console_config, fp=config_file, indent=2)
    except FileNotFoundError:
        print(f"configuration file not found: {config_path}")


if __name__ == '__main__':
    args = compose_argument_parser()

    update_config(
        config_path=args.config_path,
        ollama_ip=args.ollama_ip if args.ollama_ip != "local" else get_ip(),
        ollama_port=args.ollama_port,
    )
