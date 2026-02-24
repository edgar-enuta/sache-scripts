from abc import ABC, abstractmethod


class BaseScraper(ABC):
    name: str

    @abstractmethod
    def authenticate(self) -> bool:
        """Authenticate with the website. Returns True on success."""
        ...

    @abstractmethod
    def scrape(self, search_terms: list[str]) -> list[dict]:
        """Scrape data for the given search terms.
        Returns list of dicts, one per term, keys = Excel column names."""
        ...

    @abstractmethod
    def close(self) -> None:
        """Clean up resources (sessions, connections)."""
        ...
