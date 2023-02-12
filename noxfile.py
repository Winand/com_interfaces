import os
import tempfile
from typing import Any
import nox

locations = "com_interfaces",


def install_with_constraints(session: nox.Session, *args: str, **kwargs: Any) -> None:
    """Install packages taking into account version constraints in Poetry."""
    # NamedTemporaryFile Permission denied https://stackoverflow.com/a/54768241
    requirements = tempfile.NamedTemporaryFile(delete=False)
    try:
        session.run(
            "poetry",
            "export",
            "--with=dev",
            "--format=constraints.txt",
            "--without-hashes",
            f"--output={requirements.name}",
            external=True,
        )
        session.install(f"--constraint={requirements.name}", *args, **kwargs)
    finally:
        requirements.close()
        os.unlink(requirements.name)


@nox.session(python=["3.10", "3.9", "3.8"])
def lint(session: nox.Session):
    args = session.posargs or locations
    session.install(
        "flake8", "flake8-annotations", "flake8-black", "flake8-isort"
    )
    session.run("flake8", *args)


@nox.session(python=["3.10", "3.9", "3.8"])
def mypy(session: nox.Session) -> None:
    """Type-check using mypy."""
    args = session.posargs or locations
    install_with_constraints(session, "mypy")
    session.run("mypy", "--install-types", "--non-interactive", *args)


@nox.session(python=["3.10", "3.9", "3.8"])
def pyright(session: nox.Session) -> None:
    """Run the static type checker pyright."""
    args = session.posargs or locations
    install_with_constraints(session, "pyright")
    session.run("pyright", *args)
