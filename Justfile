# Development tasks for sqlite2xlsx

set dotenv-load

# Build the development Docker image
devbuild:
    docker build -f Dockerfile.dev -t sqlite2xlsx-dev .

# Start an interactive shell in the dev container
devshell:
    docker run -it --rm \
        -v $(pwd):/workspace \
        -v "$SSH_AGENT_SOCK:/run/ssh-agent.sock:ro" \
        -e SSH_AUTH_SOCK=/run/ssh-agent.sock \
        -w /workspace \
        sqlite2xlsx-dev
