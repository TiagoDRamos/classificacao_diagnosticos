import requests


class ICD11:
    def __init__(
        self,
        base_url="http://localhost:8000",
        release="2026-01",
        language="pt",
    ):
        self.base_url = base_url.rstrip("/")
        self.release = release
        self.language = language

        self.headers = {
            "API-Version": "v2",
            "Accept": "application/json",
            "Accept-Language": language,
        }

        self.mms_prefix = (
            f"http://id.who.int/icd/release/11/"
            f"{self.release}/mms/"
        )

        self.root_uri = (
            f"http://id.who.int/icd/release/11/"
            f"{self.release}/mms"
        )

    def _get(self, url, params=None):
        response = requests.get(
            url,
            headers=self.headers,
            params=params,
            timeout=10,
        )

        response.raise_for_status()
        return response.json()

    def search(self, query):
        """Pesquisa na CID-11 MMS."""
        url = (
            f"{self.base_url}/icd/release/11/"
            f"{self.release}/mms/search"
        )

        params = {
            "q": query,
            "highlightingEnabled": "false",
        }

        data = self._get(url, params)

        return data.get("destinationEntities", [])

    def get_entity(self, mms_uri):
        """Obtém uma entidade MMS através da URI."""

        if not mms_uri.startswith("http://"):
            mms_id = mms_uri
        else:
            if not mms_uri.startswith(self.mms_prefix):
                raise ValueError(
                    f"URI MMS inválida: {mms_uri}"
                )

            mms_id = mms_uri[len(self.mms_prefix):]

        # Entidades pós-coordenadas
        if " / " in mms_id:
            mms_id = mms_id.split(" / ")[0]

        url = (
            f"{self.base_url}/icd/release/11/"
            f"{self.release}/mms/{mms_id}"
        )

        return self._get(url)

    def get_hierarchy(self, mms_uri):
        """Percorre a hierarquia MMS até ao topo."""

        hierarchy = []
        visited = set()

        current_uri = mms_uri

        while current_uri:

            # Chegámos ao topo da MMS
            if current_uri == self.root_uri:
                break

            if current_uri in visited:
                break

            visited.add(current_uri)

            entity = self.get_entity(current_uri)

            title_data = entity.get("title", {})
            title = title_data.get("@value", "")

            hierarchy.append({
                "id": entity.get("@id"),
                "code": entity.get("code"),
                "title": title,
                "class_kind": entity.get("classKind"),
            })

            parents = entity.get("parent", [])

            if not parents:
                break

            current_uri = parents[0]

        hierarchy.reverse()

        return hierarchy

    def _is_simple_result(self, result):
        """
        Determina se o resultado é uma entidade simples
        ou uma expressão pós-coordenada.
        """

        entity_type = result.get("entityType")

        # entityType 0 corresponde normalmente a entidade simples.
        if entity_type == 0:
            return True

        # Se o ID não contém uma expressão composta,
        # também consideramos simples.
        result_id = result.get("id", "")

        return " / " not in result_id

    def _format_result(self, result):
        """Converte um resultado da pesquisa para um formato simples."""

        matching_pvs = result.get("matchingPVs", [])

        foundation_uri = None

        if matching_pvs:
            foundation_uri = matching_pvs[0].get("foundationUri")

        return {
            "title": result.get("title"),
            "code": result.get("theCode"),
            "chapter": result.get("chapter"),
            "score": result.get("score"),
            "mms_id": result.get("stemId", result.get("id")),
            "foundation_uri": foundation_uri,
            "is_leaf": result.get("isLeaf"),
            "is_simple": self._is_simple_result(result),
            "title_is_top_score": result.get("titleIsTopScore"),
            "important": result.get("important"),
        }

    def classify(
        self,
        query,
        min_score=0.80,
        ambiguity_margin=0.10,
    ):
        """
        Classifica uma descrição através da CID-11 MMS.

        Retorna:
            - match: correspondência principal
            - confidence: high / medium / low
            - ambiguous: True/False
            - alternatives: outras correspondências
            - hierarchy: hierarquia oficial MMS
        """

        results = self.search(query)

        if not results:
            return {
                "query": query,
                "match": None,
                "confidence": "none",
                "ambiguous": False,
                "alternatives": [],
                "hierarchy": [],
            }

        # Preferimos entidades simples.
        simple_results = [
            r for r in results
            if self._is_simple_result(r)
        ]

        if simple_results:
            results_to_use = simple_results
        else:
            results_to_use = results

        # A API já devolve os resultados ordenados por relevância.
        best = results_to_use[0]

        best_score = best.get("score", 0)

        second_score = 0

        if len(results_to_use) > 1:
            second_score = results_to_use[1].get("score", 0)

        # Determinar confiança
        if best_score >= min_score:
            confidence = "high"
        elif best_score >= 0.60:
            confidence = "medium"
        else:
            confidence = "low"

        # Determinar ambiguidade
        ambiguous = (
            len(results_to_use) > 1
            and (best_score - second_score) < ambiguity_margin
        )

        match = self._format_result(best)

        # Obter hierarquia apenas para o resultado escolhido.
        hierarchy = self.get_hierarchy(
            match["mms_id"]
        )

        match["hierarchy"] = hierarchy

        # Alternativas
        alternatives = []

        for result in results_to_use[1:6]:
            alternatives.append(
                self._format_result(result)
            )

        return {
            "query": query,
            "match": match,
            "confidence": confidence,
            "ambiguous": ambiguous,
            "alternatives": alternatives,
            "hierarchy": hierarchy,
        }