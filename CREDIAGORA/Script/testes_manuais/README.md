# Testes Manuais do ERCard

Estes scripts registram os testes controlados usados durante o desenvolvimento.
Nao fazem parte da automacao normal nem da suite offline em `tests/`.
Alguns representam etapas anteriores a correcoes de foco, janela e tempo de espera.
Nao devem ser usados como substitutos do fluxo atual em `start.py`.

Os imports de producao sao resolvidos por `_ambiente.py`; os modulos principais
continuam em `CREDIAGORA/Script`.

ATENCAO: executar um destes arquivos abre uma sessao real e pode exportar dados.
Nao executar enquanto outra automacao ou exportacao estiver em andamento.
Antes de usar, conferir as acoes e o ponto de parada do script escolhido.

Para uso normal, iniciar pela raiz do projeto:

```powershell
python -u .\CREDIAGORA\Script\start.py --somente-ercard
```
