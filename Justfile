# Development tasks for sqlite-to-xlsx

set dotenv-load

# List available commands
default:
    @just --list

# Build debug binary
build:
    cargo build

# Run tests
test:
    cargo test

# Build release binary
build-release:
    cargo build --release

# Build the development Docker image
devbuild:
    docker build -f Dockerfile.dev -t sqlite-to-xlsx-dev .

# Start an interactive shell in the dev container
devshell:
    docker run -it --rm \
        -v $(pwd):/workspace \
        -v "$SSH_AGENT_SOCK:/run/ssh-agent.sock:ro" \
        -e SSH_AUTH_SOCK=/run/ssh-agent.sock \
        -w /workspace \
        sqlite-to-xlsx-dev
