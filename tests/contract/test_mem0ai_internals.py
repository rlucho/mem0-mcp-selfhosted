"""Contract tests: mem0ai internal API stability.

These tests validate assumptions about mem0ai internals that our code depends on.
If these fail after a mem0ai upgrade, our code needs updating.

NOTE: These tests require mem0ai to be installed. They test the real package,
not mocks. Skip with `pytest -m "not contract"` if deps unavailable.
"""

from __future__ import annotations

import pytest

# Mark all tests in this module as contract tests
pytestmark = pytest.mark.contract


class TestVectorStoreClientAccess:
    """Test that memory.vector_store.client is a public, stable attribute."""

    def test_qdrant_class_has_client_attribute(self):
        """The Qdrant vector store class exposes .client as a public attribute."""
        try:
            from mem0.vector_stores.qdrant import Qdrant
        except ImportError:
            pytest.skip("mem0ai not installed")

        # Verify 'client' is in the class (not a private _client)
        assert hasattr(Qdrant, "__init__"), "Qdrant class must have __init__"
        # Check the source to verify client is assigned (not _client)
        import inspect

        source = inspect.getsource(Qdrant.__init__)
        assert "self.client" in source, (
            "INVARIANT BROKEN: Qdrant.__init__ must assign self.client. "
            "Our code accesses memory.vector_store.client directly."
        )

    def test_qdrant_class_has_collection_name(self):
        """The Qdrant vector store class exposes .collection_name."""
        try:
            from mem0.vector_stores.qdrant import Qdrant
        except ImportError:
            pytest.skip("mem0ai not installed")

        import inspect

        source = inspect.getsource(Qdrant.__init__)
        assert "self.collection_name" in source, (
            "INVARIANT BROKEN: Qdrant.__init__ must assign self.collection_name. "
            "Our code accesses memory.vector_store.collection_name."
        )


class TestMcpSdkImports:
    """Test MCP SDK import paths remain stable."""

    def test_mcp_client_session_importable(self):
        """ClientSession import path remains valid across MCP SDK versions."""
        try:
            from mcp.client.session import ClientSession
        except ImportError:
            pytest.skip("mcp SDK not installed")

        assert ClientSession  # Import succeeded — contract satisfied


class TestLlmFactoryRegistration:
    """Test LlmFactory.register_provider() behavior."""

    def test_register_provider_exists(self):
        """LlmFactory has a register_provider classmethod."""
        try:
            from mem0.utils.factory import LlmFactory
        except ImportError:
            pytest.skip("mem0ai not installed")

        assert hasattr(LlmFactory, "register_provider"), (
            "INVARIANT BROKEN: LlmFactory must have register_provider classmethod."
        )

    def test_register_provider_is_idempotent(self):
        """Calling register_provider twice with same name doesn't error."""
        try:
            from mem0.utils.factory import LlmFactory
        except ImportError:
            pytest.skip("mem0ai not installed")

        # Register once
        LlmFactory.register_provider(
            name="test_idempotent",
            class_path="mem0_mcp_selfhosted.llm_anthropic.AnthropicOATLLM",
            config_class=None,
        )
        # Register again — should not raise
        LlmFactory.register_provider(
            name="test_idempotent",
            class_path="mem0_mcp_selfhosted.llm_anthropic.AnthropicOATLLM",
            config_class=None,
        )

    def test_registration_persists_across_calls(self):
        """Registered provider persists in factory after registration."""
        try:
            from mem0.utils.factory import LlmFactory
        except ImportError:
            pytest.skip("mem0ai not installed")

        LlmFactory.register_provider(
            name="test_persist",
            class_path="mem0_mcp_selfhosted.llm_anthropic.AnthropicOATLLM",
            config_class=None,
        )

        # Verify the provider is in the factory's registry
        # The factory uses a class-level dict, so it should persist
        provider_map = getattr(LlmFactory, "provider_to_class", None)
        if provider_map is not None:
            assert "test_persist" in provider_map, (
                "INVARIANT BROKEN: Registered provider must persist in LlmFactory."
            )


class TestOllamaLLMInterface:
    """Validate upstream OllamaLLM interface our subclass depends on."""

    def test_ollama_llm_has_parse_response(self):
        """OllamaLLM has _parse_response method we override."""
        try:
            from mem0.llms.ollama import OllamaLLM
        except ImportError:
            pytest.skip("mem0ai not installed")

        assert hasattr(OllamaLLM, "_parse_response"), (
            "INVARIANT BROKEN: OllamaLLM must have _parse_response method. "
            "Our OllamaToolLLM subclass overrides it."
        )

    def test_ollama_llm_has_generate_response(self):
        """OllamaLLM has generate_response method we override."""
        try:
            from mem0.llms.ollama import OllamaLLM
        except ImportError:
            pytest.skip("mem0ai not installed")

        assert hasattr(OllamaLLM, "generate_response"), (
            "INVARIANT BROKEN: OllamaLLM must have generate_response method. "
            "Our OllamaToolLLM subclass overrides it."
        )

    def test_ollama_config_has_base_url(self):
        """OllamaConfig accepts ollama_base_url parameter."""
        try:
            from mem0.configs.llms.ollama import OllamaConfig
        except ImportError:
            pytest.skip("mem0ai not installed")

        # Verify __init__ accepts ollama_base_url and stores it
        cfg = OllamaConfig(ollama_base_url="http://test:11434")
        assert cfg.ollama_base_url == "http://test:11434", (
            "INVARIANT BROKEN: OllamaConfig must accept and store ollama_base_url. "
            "Our config.py passes this field to Ollama LLM config."
        )

    def test_ollama_llm_init_accepts_config(self):
        """OllamaLLM.__init__ accepts a 'config' parameter by name."""
        try:
            from mem0.llms.ollama import OllamaLLM
        except ImportError:
            pytest.skip("mem0ai not installed")

        import inspect

        sig = inspect.signature(OllamaLLM.__init__)
        params = list(sig.parameters.keys())
        assert "config" in params, (
            "INVARIANT BROKEN: OllamaLLM.__init__ must accept a 'config' parameter. "
            f"Our OllamaToolLLM inherits __init__ from it. Found params: {params}"
        )


class TestEntityStoreInternals:
    """entity_store.py swaps three Memory methods and wraps the entity_store property.

    Method bodies are read from mem0's source file, not from the live class, because the
    class may already carry our patch in this process.
    """

    REPLACED = ("_existing_entities_by_text", "_remove_memory_from_entity_store", "_link_entities_for_memory")

    @staticmethod
    def _memory_methods():
        """Map each Memory method name to (ast node, source text), parsed from mem0/memory/main.py."""
        try:
            import mem0.memory.main as main_module
        except ImportError:
            pytest.skip("mem0ai not installed")
        import ast
        import inspect

        source = inspect.getsource(main_module)
        tree = ast.parse(source)
        memory_cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == "Memory")
        return {
            n.name: (n, ast.get_source_segment(source, n))
            for n in memory_cls.body
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))
        }

    def test_replaced_methods_keep_their_signatures(self):
        import inspect

        from mem0_mcp_selfhosted import entity_store

        methods = self._memory_methods()
        for name in self.REPLACED:
            assert name in methods, f"INVARIANT BROKEN: Memory.{name} no longer exists. entity_store.py replaces it."
            upstream = [a.arg for a in methods[name][0].args.args]
            ours = list(inspect.signature(getattr(entity_store, name.lstrip("_"))).parameters)
            assert upstream == ours, (
                f"INVARIANT BROKEN: Memory.{name}{tuple(upstream)} no longer matches "
                f"entity_store.{name.lstrip('_')}{tuple(ours)}."
            )

    def test_entity_store_is_a_lazy_property_over_a_private_attribute(self):
        import ast

        methods = self._memory_methods()
        node, source = methods["entity_store"]
        assert any(isinstance(d, ast.Name) and d.id == "property" for d in node.decorator_list), (
            "INVARIANT BROKEN: Memory.entity_store must be a property. entity_store.py wraps its getter."
        )
        assert "self._entity_store" in source
        assert "self._entity_store = None" in methods["__init__"][1]

    def test_upstream_still_lists_a_capped_page_of_entities(self):
        """The reason entity_store.py exists. If this fails, check whether the patch is still needed."""
        methods = self._memory_methods()
        for name in ("_existing_entities_by_text", "_remove_memory_from_entity_store"):
            assert "top_k=10000" in methods[name][1], (
                f"Memory.{name} no longer lists top_k=10000 entities. Check whether mem0ai fixed "
                "the entity lookup upstream before keeping entity_store.py."
            )

    def test_entity_lookups_are_only_read_with_get(self):
        """existing_entities_by_text() returns an object that implements .get() and nothing else."""
        import re

        methods = self._memory_methods()
        callers = sorted(
            name
            for name, (_, source) in methods.items()
            if name != "_existing_entities_by_text" and "_existing_entities_by_text(" in source
        )
        assert callers == ["_add_to_vector_store", "_upsert_entity"], (
            f"INVARIANT BROKEN: callers of _existing_entities_by_text changed: {callers}. "
            "Check they only call .get() on the result."
        )
        assert "self._existing_entities_by_text(search_filters).get(" in methods["_upsert_entity"][1]
        uses = re.findall(
            r"\bexact_matches\b(\s*=\s*self\._existing_entities_by_text\(|\.get\()?",
            methods["_add_to_vector_store"][1],
        )
        assert uses and all(uses), "INVARIANT BROKEN: _add_to_vector_store reads exact_matches other than via .get()."

    def test_update_and_delete_clean_up_through_the_replaced_methods(self):
        methods = self._memory_methods()
        update_source = methods["_update_memory"][1]
        assert "self._remove_memory_from_entity_store(" in update_source
        assert "self._link_entities_for_memory(" in update_source
        assert "self._remove_memory_from_entity_store(" in methods["_delete_memory"][1]

    def test_entity_payload_fields_and_scope_keys(self):
        """Lookups and maintenance read data and linked_memory_ids, scoped by user_id, agent_id, run_id."""
        methods = self._memory_methods()
        add_source = methods["_add_to_vector_store"][1]
        for fragment in ('"data": entity_text', '"linked_memory_ids": sorted(memory_ids)', "**search_filters"):
            assert fragment in add_source, f"INVARIANT BROKEN: the entity insert payload lost {fragment!r}"
        assert '("user_id", "agent_id", "run_id")' in methods["_delete_memory"][1]

    def test_semantic_dedup_threshold_matches(self):
        import re

        from mem0_mcp_selfhosted import entity_store

        upsert_source = self._memory_methods()["_upsert_entity"][1]
        assert {float(t) for t in re.findall(r"score >= ([0-9.]+)", upsert_source)} == {entity_store.SEMANTIC_MATCH_SCORE}

    def test_normalize_entity_text_matches_mem0(self):
        try:
            from mem0.memory.main import Memory
        except ImportError:
            pytest.skip("mem0ai not installed")
        from mem0_mcp_selfhosted.entity_store import normalize_entity_text

        for sample in ("Proxmox VE", "  Samsung\t990  PRO ", "École Normale", "a b", ""):
            assert normalize_entity_text(sample) == Memory._normalize_entity_text(sample)

    def test_qdrant_wrapper_surface_used_by_entity_store(self):
        try:
            import mem0.memory.main as main_module
            from mem0.utils.entity_extraction import extract_entities
            from mem0.vector_stores.qdrant import Qdrant
        except ImportError:
            pytest.skip("mem0ai not installed")
        import inspect

        assert "set_payload" in inspect.getsource(Qdrant.update), (
            "INVARIANT BROKEN: Qdrant.update(vector=None, payload=...) must merge the payload via set_payload."
        )
        assert "self.is_local" in inspect.getsource(Qdrant.__init__)
        assert main_module._entity_collection_name("qdrant", "abc") == "abc_entities"
        assert main_module.extract_entities is extract_entities
