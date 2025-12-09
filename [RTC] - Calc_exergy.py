import os
import win32com.client as win32

class AspenAnalyzer:
    def __init__(self):
        self.aspen = None
        self.results = {}
        self.T0 = 298.15  # Temperatura ambiente de referência (K)

    def connect_to_aspen(self, file_path):
        """Estabelece conexão com o arquivo Aspen Plus"""
        try:
            self.aspen = win32.Dispatch("Apwn.Document")
            self.aspen.InitFromArchive2(os.path.abspath(file_path))
            print("Conexão com Aspen Plus estabelecida com sucesso!")

            self.run_simulation()
            return True
        except Exception as e:
            print(f"Erro ao conectar ao Aspen Plus: {e}")
            return False

    def run_simulation(self):
        """Executa a simulação do Aspen Plus"""
        try:
            print("Executando simulação Aspen Plus...")
            self.aspen.Engine.Run2()
            print("Simulação executada com sucesso!")
        except Exception as e:
            print(f"Simulação já executada ou erro: {e}")

    def close_connection(self):
        """Fecha a conexão com o Aspen Plus"""
        if self.aspen:
            self.aspen.Close()
            print("Conexão com Aspen Plus encerrada.")

    def get_node_value(self, node_path, default=0.0):
        """Obtém o valor de um nó com tratamento de erros"""
        try:
            node = self.aspen.Tree.FindNode(node_path)
            if node is None:
                print(f"AVISO: Nó não encontrado: {node_path}")
                return default

            # Verificar se o nó possui valor válido
            if hasattr(node, 'Value') and node.Value is not None:
                value = float(node.Value)
                return value
            else:
                print(f"AVISO: Nó sem valor válido: {node_path}")
                return default

        except Exception as e:
            print(f"ERRO ao acessar nó {node_path}: {e}")
            return default

    # ==============================================
    # SISTEMA PADRONIZADO DE OBTENÇÃO DE VALORES
    # ==============================================

    def get_stream_exergy(self, stream_name, default=0.0):
        """Obtém a exergia de uma corrente"""
        path = f"\\Data\\Streams\\{stream_name}\\Output\\STRM_UPP\\EXERGYFL\\MIXED\\TOTAL"
        return self.get_node_value(path, default)

    def get_equipment_power(self, equipment_name, default=0.0):
        """Obtém a potência de um equipamento"""
        path = f"\\Data\\Blocks\\{equipment_name}\\Output\\WNET"
        return self.get_node_value(path, default)

    def get_heat_duty(self, equipment_name, default=0.0):
        """Obtém o calor trocado em equipamentos"""
        path = f"\\Data\\Blocks\\{equipment_name}\\Output\\QCALC"
        return self.get_node_value(path, default)

    def get_reboiler_duty(self, column_name, default=0.0):
        """Obtém o calor do reboiler de uma coluna"""
        path = f"\\Data\\Blocks\\{column_name}\\Output\\REB_DUTY"
        return self.get_node_value(path, default)

    def get_condenser_duty(self, column_name, default=0.0):
        """Obtém o calor do condensador de uma coluna"""
        path = f"\\Data\\Blocks\\{column_name}\\Output\\COND_DUTY"
        return self.get_node_value(path, default)

    def get_flash_heat_duty(self, flash_name, default=0.0):
        """Obtém o calor trocado em tanques flash"""
        path = f"\\Data\\Blocks\\{flash_name}\\Output\\QCALC"
        return self.get_node_value(path, default)

    # ==============================================
    # CÁLCULOS DE EXERGIA TÉRMICA
    # ==============================================

    def calculate_exergy_heat_cooler(self, heat_duty):
        """Calcula a exergia associada a uma transferência de calor para resfriadores"""
        try:
            if heat_duty is None or heat_duty == 0:
                return 0.0
            temp_cooler = 303.15
            # Usar valor absoluto para cálculo da exergia
            return abs(heat_duty) * (1 - self.T0 / temp_cooler)
        except:
            return 0.0

    def calculate_exergy_heat_furnace(self, heat_duty):
        """Calcula a exergia associada a uma transferência de calor para fornos"""
        try:
            if heat_duty is None or heat_duty == 0:
                return 0.0
            temp_furnace = 3273.15
            # Usar valor absoluto para cálculo da exergia
            return abs(heat_duty) * (1 - self.T0 / temp_furnace)
        except:
            return 0.0

    def calculate_exergy_heat_flash(self, heat_duty):
        """Calcula a exergia associada a uma transferência de calor para tanques flash"""
        try:
            if heat_duty is None or heat_duty == 0:
                return 0.0
            temp_flash = 313.15
            # Usar valor absoluto para cálculo da exergia
            return abs(heat_duty) * (1 - self.T0 / temp_flash)
        except:
            return 0.0

    def calculate_exergy_heat_reactor(self, heat_duty):
        """Calcula a exergia associada a uma transferência de calor para reatores"""
        try:
            if heat_duty is None or heat_duty == 0:
                return 0.0
            temp_reactor = 303.15
            # Usar valor absoluto para cálculo da exergia
            return abs(heat_duty) * (1 - self.T0 / temp_reactor)
        except:
            return 0.0

    def calculate_exergy_heat_condenser(self, heat_duty):
        """Calcula a exergia associada a uma transferência de calor para condensadores"""
        try:
            if heat_duty is None or heat_duty == 0:
                return 0.0
            temp_condenser = 333.15
            # Usar valor absoluto para cálculo da exergia
            return abs(heat_duty) * (1 - self.T0 / temp_condenser)
        except:
            return 0.0

    def calculate_exergy_heat_reboiler(self, heat_duty):
        """Calcula a exergia associada a uma transferência de calor para reboilers"""
        try:
            if heat_duty is None or heat_duty == 0:
                return 0.0
            temp_reboiler = 570.15
            # Usar valor absoluto para cálculo da exergia
            return abs(heat_duty) * (1 - self.T0 / temp_reboiler)
        except:
            return 0.0
        
    def calculate_exergy_heat_compressor(self, heat_duty):
        """Calcula a exergia associada a uma transferência de calor para compressores"""
        try:
            if heat_duty is None or heat_duty == 0:
                return 0.0
            temp_compressor = 350.15  # Temperatura típica de compressores
            return abs(heat_duty) * (1 - self.T0 / temp_compressor)
        except:
            return 0.0

    # ==============================================
    # ANÁLISES POR TIPO DE EQUIPAMENTO
    # ==============================================

    def calculate_pumps_exergy_loss(self):
        """Calcula perda exergética em bombas"""
        pumps = [
            {"name": "PUMP-1", "input": "TGO-1", "output": "TGO-2"},
            {"name": "PUMP-2", "input": "ALKENE6", "output": "ALKENE7"},
            {"name": "PUMP-3", "input": "H-DIESEL", "output": "H2DIESEL"},
            {"name": "PUMP-4", "input": "HOT-AK15", "output": "ALKENE16"}
        ]

        total_loss = 0.0
        print("\n" + "="*50)
        print("ANÁLISE EXERGÉTICA DE BOMBAS")
        print("="*50)

        for pump in pumps:
            try:
                input_ex = self.get_stream_exergy(pump['input'])
                output_ex = self.get_stream_exergy(pump['output'])
                power = self.get_equipment_power(pump['name'])

                loss = input_ex + power - output_ex
                efficiency = (1 - (loss / (input_ex + power))) * 100 if (input_ex + power) > 0 else 0

                print(f"\nBomba {pump['name']}:")
                print(f"  Entrada ({pump['input']}): {input_ex:.2f} kW")
                print(f"  Saída ({pump['output']}): {output_ex:.2f} kW")
                print(f"  Potência: {power:.2f} kW")
                print(f"  Perda exergética: {loss:.2f} kW")
                print(f"  Eficiência: {efficiency:.2f}%")

                total_loss += max(loss, 0)

            except Exception as e:
                print(f"Erro na análise da bomba {pump['name']}: {e}")
                continue

        print(f"\nTotal de perda em bombas: {total_loss:.2f} kW")
        self.results['bombas'] = total_loss
        return total_loss

    def calculate_compressors_exergy_loss(self):
        """Calcula perda exergética em compressores"""
        compressors = [
            {"name": "COMPR-1", "input": "GASES4", "output": "GASES-5", "type": "standard"},
            {"name": "M-COMPR", "input": "H2-REC-2", "output": "M-H2-REC", "type": "m-compressor"},
            {"name": "M-COMPR2", "input": "H2-REC-3", "output": "H2-REC-4", "type": "m-compressor"}
        ]

        total_loss = 0.0
        print("\n" + "="*50)
        print("ANÁLISE EXERGÉTICA DE COMPRESSORES")
        print("="*50)

        for comp in compressors:
            try:
                input_ex = self.get_stream_exergy(comp['input'])
                output_ex = self.get_stream_exergy(comp['output'])
                
                if comp['type'] == "standard":
                    # Compressores padrão - apenas potência
                    power = self.get_equipment_power(comp['name'])
                    heat_duty = 0.0
                    exergy_heat = 0.0
                    loss = input_ex + power - output_ex
                    
                    print(f"\nCompressor {comp['name']} (Standard):")
                    print(f"  Entrada ({comp['input']}): {input_ex:.2f} kW")
                    print(f"  Saída ({comp['output']}): {output_ex:.2f} kW")
                    print(f"  Potência: {power:.2f} kW")
                    print(f"  Calor trocado: {heat_duty:.2f} kW")
                    
                else:
                    # Compressores M-COMPR - potência e calor
                    power = self.get_equipment_power(comp['name'])
                    heat_duty = self.get_heat_duty(comp['name'])
                    exergy_heat = self.calculate_exergy_heat_compressor(heat_duty)
                    
                    # Para compressores, o calor é geralmente removido (negativo)
                    if heat_duty < 0:
                        # Calor removido: não é entrada, mas sim uma saída de exergia
                        loss = input_ex + power - output_ex - exergy_heat
                    else:
                        # Calor fornecido: entrada de exergia
                        loss = input_ex + power + exergy_heat - output_ex
                    
                    print(f"\nCompressor {comp['name']} (M-COMPR):")
                    print(f"  Entrada ({comp['input']}): {input_ex:.2f} kW")
                    print(f"  Saída ({comp['output']}): {output_ex:.2f} kW")
                    print(f"  Potência: {power:.2f} kW")
                    print(f"  Calor trocado: {heat_duty:.2f} kW")
                    print(f"  Exergia do calor: {exergy_heat:.2f} kW")
                    if heat_duty < 0:
                        print("  → Calor REMOVIDO (saída do sistema)")
                    else:
                        print("  → Calor FORNECIDO (entrada no sistema)")

                efficiency = (1 - (loss / (input_ex + power))) * 100 if (input_ex + power) > 0 else 0

                print(f"  Perda exergética: {loss:.2f} kW")
                print(f"  Eficiência: {efficiency:.2f}%")

                total_loss += max(loss, 0)

            except Exception as e:
                print(f"Erro na análise do compressor {comp['name']}: {e}")
                continue

        print(f"\nTotal de perda em compressores: {total_loss:.2f} kW")
        self.results['compressores'] = total_loss
        return total_loss

    def calculate_coolers_exergy_loss(self):
        """Calcula perda exergética em resfriadores"""
        coolers = [
            {"name": "COOLER-1", "input": "TGO+H2-5", "output": "ALKENE1"},
            {"name": "COOLER-1", "input": "ALKENE3", "output": "ALKENE4"},
            {"name": "COOLER-2", "input": "ALKENE12", "output": "ALKENE13"},
            {"name": "COOLER-4", "input": "C-DIESEL", "output": "DIESEL"},
            {"name": "COOLER-5", "input": "QBIO-QAV", "output": "BIO-QAV"}
        ]

        total_loss = 0.0
        print("\n" + "="*50)
        print("ANÁLISE EXERGÉTICA DE RESFRIADORES")
        print("="*50)

        for cooler in coolers:
            try:
                input_ex = self.get_stream_exergy(cooler['input'])
                output_ex = self.get_stream_exergy(cooler['output'])
                heat_duty = self.get_heat_duty(cooler['name'])

                # Para resfriadores, o calor é removido (negativo)
                exergy_heat = self.calculate_exergy_heat_cooler(heat_duty)
                
                # Balanço: entrada = saída + perda
                # Calor removido é uma saída de exergia
                loss = input_ex - output_ex - exergy_heat
                efficiency = (1 - (loss / input_ex)) * 100 if input_ex > 0 else 0

                print(f"\nResfriador {cooler['name']}:")
                print(f"  Entrada ({cooler['input']}): {input_ex:.2f} kW")
                print(f"  Saída ({cooler['output']}): {output_ex:.2f} kW")
                print(f"  Calor removido: {heat_duty:.2f} kW")
                print(f"  Exergia do calor: {exergy_heat:.2f} kW")
                print(f"  Perda exergética: {loss:.2f} kW")
                print(f"  Eficiência: {efficiency:.2f}%")

                total_loss += max(loss, 0)

            except Exception as e:
                print(f"Erro na análise do resfriador {cooler['name']}: {e}")
                continue

        print(f"\nTotal de perda em resfriadores: {total_loss:.2f} kW")
        self.results['resfriadores'] = total_loss
        return total_loss

    def calculate_mixers_exergy_loss(self):
        """Calcula perda exergética em misturadores"""
        mixers = [
            {"name": "MIX-1", "inputs": ["TGO-2", "H2-TO-R1"], "output": "TGO+H2"},
            {"name": "MIXER-2", "inputs": ["GASES1","GASES2", "GASES3"], "output": "GASES4"},
            {"name": "MIXER-3", "inputs": ["H2-REC-4", "ALKENE7", "MKUP-R3"], "output": "ALKENE8"},
            {"name": "MIX-4", "inputs": ["MKUP-R1", "M-H2-REC"], "output": "H2-TO-R1"}
        ]

        total_loss = 0.0
        print("\n" + "="*50)
        print("ANÁLISE EXERGÉTICA DE MISTURADORES")
        print("="*50)

        for mixer in mixers:
            try:
                print(f"\nMisturador {mixer['name']}:")

                input_ex = 0.0
                for stream in mixer['inputs']:
                    stream_ex = self.get_stream_exergy(stream)
                    print(f"  Entrada ({stream}): {stream_ex:.2f} kW")
                    input_ex += stream_ex

                output_ex = self.get_stream_exergy(mixer['output'])
                print(f"  Saída ({mixer['output']}): {output_ex:.2f} kW")

                loss = input_ex - output_ex
                efficiency = (1 - (loss / input_ex)) * 100 if input_ex > 0 else 0

                print(f"  Perda exergética: {loss:.2f} kW")
                print(f"  Eficiência: {efficiency:.2f}%")
                total_loss += max(loss, 0)

            except Exception as e:
                print(f"Erro na análise do misturador {mixer['name']}: {e}")
                continue

        print(f"\nTotal de perda em misturadores: {total_loss:.2f} kW")
        self.results['misturadores'] = total_loss
        return total_loss

    def calculate_valves_exergy_loss(self):
        """Calcula perda exergética em válvulas"""
        valves = [
            {"name": "VALVE-1", "input": "ALKENE4", "output": "ALKENE5"},
            {"name": "VALVE-2", "input": "ALKENE14", "output": "ALKENE15"}
        ]

        total_loss = 0.0
        print("\n" + "="*50)
        print("ANÁLISE EXERGÉTICA DE VÁLVULAS")
        print("="*50)

        for valve in valves:
            try:
                input_ex = self.get_stream_exergy(valve['input'])
                output_ex = self.get_stream_exergy(valve['output'])

                loss = input_ex - output_ex
                efficiency = (1 - (loss / input_ex)) * 100 if input_ex > 0 else 0

                print(f"\nVálvula {valve['name']}:")
                print(f"  Entrada ({valve['input']}): {input_ex:.2f} kW")
                print(f"  Saída ({valve['output']}): {output_ex:.2f} kW")
                print(f"  Perda exergética: {loss:.2f} kW")
                print(f"  Eficiência: {efficiency:.2f}%")

                total_loss += max(loss, 0)

            except Exception as e:
                print(f"Erro na análise da válvula {valve['name']}: {e}")
                continue

        print(f"\nTotal de perda em válvulas: {total_loss:.2f} kW")
        self.results['valvulas'] = total_loss
        return total_loss

    def calculate_separators_exergy_loss(self):
        """Calcula perda exergética em separadores"""
        separators = [
            {"name": "SEP", "input": "GASES-5", "outputs": ["TAIL-GAS", "H2-REC"]}
        ]

        total_loss = 0.0
        print("\n" + "="*50)
        print("ANALISE EXERGETICA DE SEPARADORES")
        print("="*50)

        for separator in separators:
            try:
                print(f"\nSeparador {separator['name']}:")

                input_ex = self.get_stream_exergy(separator['input'])
                print(f"  Entrada ({separator['input']}): {input_ex:.2f} kW")

                output_ex = 0.0
                for stream in separator['outputs']:
                    stream_ex = self.get_stream_exergy(stream)
                    print(f"  Saída ({stream}): {stream_ex:.2f} kW")
                    output_ex += stream_ex

                loss = input_ex - output_ex
                efficiency = (1 - (loss / input_ex)) * 100 if input_ex > 0 else 0

                print(f"  Perda exergetica: {loss:.2f} kW")
                print(f"  Eficiencia: {efficiency:.2f}%")
                total_loss += max(loss, 0)

            except Exception as e:
                print(f"Erro na analise do separador {separator['name']}: {e}")
                continue

        print(f"\nTotal de perda em separadores: {total_loss:.2f} kW")
        self.results['separadores'] = total_loss
        return total_loss

    def calculate_furnaces_exergy_loss(self):
        """Calcula perda exergética em fornos"""
        furnaces = [
            {"name": "FURNACE1", "input": "TGO+H2-3", "output": "TGO+H2-4"},
            {"name": "FURNACE2", "input": "HOT-ALK9", "output": "ALKENE10"}
            ]

        total_loss = 0.0
        print("\n" + "="*50)
        print("ANÁLISE EXERGÉTICA DE FORNOS")
        print("="*50)

        for furnace in furnaces:
            try:
                input_ex = self.get_stream_exergy(furnace['input'])
                output_ex = self.get_stream_exergy(furnace['output'])
                heat_supplied = self.get_heat_duty(furnace['name'])

                exergy_heat = self.calculate_exergy_heat_furnace(heat_supplied)
                loss = input_ex + exergy_heat - output_ex
                efficiency = (1 - (loss / (input_ex + exergy_heat))) * 100 if (input_ex + exergy_heat) > 0 else 0

                print(f"\nForno {furnace['name']}:")
                print(f"  Entrada ({furnace['input']}): {input_ex:.2f} kW")
                print(f"  Saída ({furnace['output']}): {output_ex:.2f} kW")
                print(f"  Calor fornecido: {heat_supplied:.2f} kW")
                print(f"  Exergia do calor: {exergy_heat:.2f} kW")
                print(f"  Perda exergética: {loss:.2f} kW")
                print(f"  Eficiência: {efficiency:.2f}%")

                total_loss += max(loss, 0)

            except Exception as e:
                print(f"Erro na análise do forno {furnace['name']}: {e}")
                continue

        print(f"\nTotal de perda em fornos: {total_loss:.2f} kW")
        self.results['fornos'] = total_loss
        return total_loss

    def calculate_heat_exchanger_exergy_loss(self):
        """Calcula perda exergética em trocador de calor HEAT-X"""
        heat_exchangers = [
            {"name": "HEAT-1", "inputs": ["H-AK1-3", "TGO+H2"], "outputs": ["TGO+H2-1", "ALKENE2"]},
            {"name": "HEAT-2", "inputs": ["TGO+H2-1", "H2DIESEL"], "outputs": ["C-DIESEL", "TGO+H2-2"]},
            {"name": "HEAT-3", "inputs": ["TGO+H2-2", "ALKENE1"], "outputs": ["TGO+H2-3", "HOT-AK1"]},
            {"name": "HEAT-4", "inputs": ["HOT-AK1", "ALKENE15"], "outputs": ["H-AK1-2", "HOT-AK15"]},
            {"name": "HEAT-5", "inputs": ["H-AK1-2", "ALKENE8"], "outputs": ["H-AK1-3", "ALKENE9"]},
            {"name": "HEAT-6", "inputs": ["ALKENE9", "ALKENE11"], "outputs": ["HOT-ALK9", "ALKENE12"]}
        ]

        total_loss = 0.0
        print("\n" + "="*50)
        print("ANÁLISE EXERGÉTICA DE TROCADOR DE CALOR")
        print("="*50)

        for exchanger in heat_exchangers:
            try:
                print(f"\nTrocador de Calor {exchanger['name']}:")

                input_ex = 0.0
                for stream in exchanger['inputs']:
                    stream_ex = self.get_stream_exergy(stream)
                    print(f"  Entrada ({stream}): {stream_ex:.2f} kW")
                    input_ex += stream_ex

                output_ex = 0.0
                for stream in exchanger['outputs']:
                    stream_ex = self.get_stream_exergy(stream)
                    print(f"  Saída ({stream}): {stream_ex:.2f} kW")
                    output_ex += stream_ex

                loss = input_ex - output_ex
                efficiency = (1 - (loss / input_ex)) * 100 if input_ex > 0 else 0

                print(f"  Perda exergética: {loss:.2f} kW")
                print(f"  Eficiência: {efficiency:.2f}%")
                total_loss += max(loss, 0)

            except Exception as e:
                print(f"Erro na análise do trocador de calor {exchanger['name']}: {e}")
                continue

        print(f"\nTotal de perda em trocador de calor: {total_loss:.2f} kW")
        self.results['trocador_calor'] = total_loss
        return total_loss

    def calculate_flash_tanks_exergy_loss(self):
        """Calcula perda exergética em tanques flash"""
        flash_tanks = [
            {"name": "FLASH-1", "input": "ALKENE3", "outputs": ["ALKENE4", "WATER-1", "GASES1"]},
            {"name": "FLASH2", "input": "ALKENE5", "outputs": ["GASES2", "ALKENE6"]},
            {"name": "FLASH3", "input": "ALKENE13", "outputs": ["GASES3", "ALKENE14"]}
        ]

        total_loss = 0.0
        print("\n" + "="*50)
        print("ANÁLISE EXERGÉTICA DE TANQUES FLASH")
        print("="*50)

        for tank in flash_tanks:
            try:
                input_ex = self.get_stream_exergy(tank['input'])
                heat_duty = self.get_flash_heat_duty(tank['name'])
                exergy_heat = self.calculate_exergy_heat_flash(heat_duty)

                output_ex = 0.0
                for stream in tank['outputs']:
                    stream_ex = self.get_stream_exergy(stream)
                    output_ex += stream_ex

                loss = input_ex + exergy_heat - output_ex
                efficiency = (1 - (loss / (input_ex + exergy_heat))) * 100 if (input_ex + exergy_heat) > 0 else 0

                print(f"\nTanque Flash {tank['name']}:")
                print(f"  Entrada ({tank['input']}): {input_ex:.2f} kW")
                print(f"  Calor trocado: {heat_duty:.2f} kW")
                print(f"  Exergia do calor: {exergy_heat:.2f} kW")

                for stream in tank['outputs']:
                    stream_ex = self.get_stream_exergy(stream)
                    print(f"  Saída ({stream}): {stream_ex:.2f} kW")

                print(f"  Perda exergética: {loss:.2f} kW")
                print(f"  Eficiência: {efficiency:.2f}%")

                total_loss += max(loss, 0)

            except Exception as e:
                print(f"Erro na análise do tanque flash {tank['name']}: {e}")
                continue

        print(f"\nTotal de perda em tanques flash: {total_loss:.2f} kW")
        self.results['tanques_flash'] = total_loss
        return total_loss

    def calculate_columns_exergy_loss(self):
        """Calcula perda exergética em colunas de destilação"""
        columns = [
            {"name": "DEST-COL", "input": "ALKENE16", "outputs": ["LIGHTS", "QBIO-QAV", "H-DIESEL"]}
        ]

        total_loss = 0.0
        print("\n" + "="*50)
        print("ANÁLISE EXERGÉTICA DE COLUNAS")
        print("="*50)

        for column in columns:
            try:
                print(f"\nColuna {column['name']}:")

                input_ex = self.get_stream_exergy(column['input'])
                print(f"  Entrada ({column['input']}): {input_ex:.2f} kW")

                output_ex = 0.0
                for stream in column['outputs']:
                    stream_ex = self.get_stream_exergy(stream)
                    print(f"  Saída ({stream}): {stream_ex:.2f} kW")
                    output_ex += stream_ex

                reboiler_duty = self.get_reboiler_duty(column['name'])
                condenser_duty = self.get_condenser_duty(column['name'])

                exergy_reboiler = self.calculate_exergy_heat_reboiler(reboiler_duty)
                exergy_condenser = self.calculate_exergy_heat_condenser(condenser_duty)

                total_exergy_heat = exergy_reboiler + exergy_condenser

                print(f"  Reboiler: {reboiler_duty:.2f} kW")
                print(f"  Exergia do reboiler: {exergy_reboiler:.2f} kW")
                print(f"  Condensador: {condenser_duty:.2f} kW")
                print(f"  Exergia do condensador: {exergy_condenser:.2f} kW")
                print(f"  Exergia térmica líquida: {total_exergy_heat:.2f} kW")

                loss = input_ex + total_exergy_heat - output_ex
                efficiency = (1 - (loss / (input_ex + total_exergy_heat))) * 100 if (input_ex + total_exergy_heat) > 0 else 0

                print(f"  Perda exergética: {loss:.2f} kW")
                print(f"  Eficiência: {efficiency:.2f}%")
                total_loss += max(loss, 0)

            except Exception as e:
                print(f"Erro na análise da coluna {column['name']}: {e}")
                continue

        print(f"\nTotal de perda em colunas: {total_loss:.2f} kW")
        self.results['colunas'] = total_loss
        return total_loss

    def calculate_reactors_exergy_loss(self):
        """Calcula perda exergética em reatores"""
        reactors = [
            {"name": "R-1", "inputs": ["TGO+H2-4"], "outputs": ["TGO+H2-5"]},
            {"name": "R-2", "inputs": ["TGO+H2-5"], "outputs": ["ALKENE1"]},
            {"name": "R-3", "inputs": ["ALKENE10"], "outputs": ["ALKENE11"]}
        ]

        total_loss = 0.0
        print("\n" + "="*50)
        print("ANÁLISE EXERGÉTICA DE REATORES")
        print("="*50)

        for reactor in reactors:
            try:
                print(f"\nReator {reactor['name']}:")

                input_ex = 0.0
                for stream in reactor['inputs']:
                    stream_ex = self.get_stream_exergy(stream)
                    print(f"  Entrada ({stream}): {stream_ex:.2f} kW")
                    input_ex += stream_ex

                output_ex = 0.0
                for stream in reactor['outputs']:
                    stream_ex = self.get_stream_exergy(stream)
                    print(f"  Saída ({stream}): {stream_ex:.2f} kW")
                    output_ex += stream_ex

                heat_reaction = self.get_heat_duty(reactor['name'])
                exergy_heat = self.calculate_exergy_heat_reactor(heat_reaction)

                print(f"  Calor de reação: {heat_reaction:.2f} kW")
                print(f"  Exergia do calor: {exergy_heat:.2f} kW")

                loss = input_ex - output_ex - exergy_heat
                efficiency = (1 - (loss / (input_ex - exergy_heat))) * 100 if (input_ex - exergy_heat) > 0 else 0

                print(f"  Perda exergética: {loss:.2f} kW")
                print(f"  Eficiência: {efficiency:.2f}%")
                total_loss += max(loss, 0)

            except Exception as e:
                print(f"Erro na análise do reator {reactor['name']}: {e}")
                continue

        print(f"\nTotal de perda em reatores: {total_loss:.2f} kW")
        self.results['reatores'] = total_loss
        return total_loss

    # ==============================================
    # SISTEMA CORRIGIDO DE CÁLCULO DE EXERGIA TOTAL
    # ==============================================

    def calculate_total_work_and_heat_exergy(self):
        """Calcula o total de exergia de trabalho e calor fornecidos à planta"""
        total_work_exergy = 0.0
        total_heat_exergy_input = 0.0  # Exergia térmica que ENTRA no sistema
        total_heat_exergy_output = 0.0  # Exergia térmica que SAI do sistema

        print("\n" + "="*60)
        print("CÁLCULO DE EXERGIA DE TRABALHO E CALOR")
        print("="*60)

        # Exergia de trabalho (bombas e compressores) - SEMPRE ENTRADA
        print("\nEXERGIA DE TRABALHO (ENTRADA):")

        # Bombas
        pumps = ["PUMP-1", "PUMP-2", "PUMP-3", "PUMP-4"]
        for pump in pumps:
            power = self.get_equipment_power(pump)
            total_work_exergy += power
            print(f"  {pump}: {power:.2f} kW")

        # Compressores
        compressors = ["COMPR-1", "M-COMPR", "M-COMPR2"]
        for comp in compressors:
            power = self.get_equipment_power(comp)
            total_work_exergy += power
            print(f"  {comp}: {power:.2f} kW")

        print(f"TOTAL EXERGIA DE TRABALHO: {total_work_exergy:.2f} kW")

        # Exergia de calor - SEPARAR ENTRE ENTRADA E SAÍDA
        print("\nEXERGIA DE CALOR:")

        # EXERGIA TÉRMICA DE ENTRADA (calor fornecido ao sistema)
        print("\n  EXERGIA TÉRMICA DE ENTRADA:")
        furnaces = ["FURNACE1", "FURNACE2"]
        for furnace in furnaces:
            heat_duty = self.get_heat_duty(furnace)
            if heat_duty > 0:  # Calor fornecido ao sistema
                exergy_heat = self.calculate_exergy_heat_furnace(heat_duty)
                total_heat_exergy_input += exergy_heat
                print(f"    {furnace}: {exergy_heat:.2f} kW (Calor: {heat_duty:.2f} kW)")

        # Reboiler da coluna - CALOR FORNECIDO (ENTRADA)
        reboiler_duty = self.get_reboiler_duty("DEST-COL")
        if reboiler_duty > 0:
            exergy_reboiler = self.calculate_exergy_heat_reboiler(reboiler_duty)
            total_heat_exergy_input += exergy_reboiler
            print(f"    Reboiler DEST-COL: {exergy_reboiler:.2f} kW (Calor: {reboiler_duty:.2f} kW)")

        # Tanques flash - CALOR FORNECIDO (ENTRADA)
        flash_tanks = ["FLASH-1", "FLASH2", "FLASH3"]
        for flash in flash_tanks:
            heat_duty = self.get_flash_heat_duty(flash)
            if heat_duty > 0:
                exergy_heat = self.calculate_exergy_heat_flash(heat_duty)
                total_heat_exergy_input += exergy_heat
                print(f"    {flash}: {exergy_heat:.2f} kW (Calor: {heat_duty:.2f} kW)")

        # EXERGIA TÉRMICA DE SAÍDA (calor removido do sistema)
        print("\n  EXERGIA TÉRMICA DE SAÍDA:")
        
        # Compressores M-COMPR - calor REMOVIDO (SAÍDA) - CORREÇÃO APLICADA
        m_compressors = ["M-COMPR", "M-COMPR2"]
        for comp in m_compressors:
            heat_duty = self.get_heat_duty(comp)
            if heat_duty < 0:  # Calor removido do compressor (SAÍDA)
                exergy_heat = self.calculate_exergy_heat_compressor(heat_duty)
                total_heat_exergy_output += exergy_heat
                print(f"    {comp}: {exergy_heat:.2f} kW (Calor REMOVIDO: {heat_duty:.2f} kW)")

        # Reatores - calor REMOVIDO (SAÍDA)
        reactors = ["R-1", "R-2", "R-3"]
        for reactor in reactors:
            heat_duty = self.get_heat_duty(reactor)
            if heat_duty < 0:  # Calor removido do reator (SAÍDA)
                exergy_heat = self.calculate_exergy_heat_reactor(heat_duty)
                total_heat_exergy_output += exergy_heat
                print(f"    {reactor}: {exergy_heat:.2f} kW (Calor REMOVIDO: {heat_duty:.2f} kW)")

        # Resfriadores - CALOR REMOVIDO (SAÍDA)
        coolers = ["COOLER-1", "COOLER-2", "COOLER-3", "COOLER-4"]
        for cooler in coolers:
            heat_duty = self.get_heat_duty(cooler)
            if heat_duty < 0:  # Calor removido do sistema
                exergy_heat = self.calculate_exergy_heat_cooler(heat_duty)
                total_heat_exergy_output += exergy_heat
                print(f"    {cooler}: {exergy_heat:.2f} kW (Calor: {heat_duty:.2f} kW)")

        # Condensador da coluna - CALOR REMOVIDO (SAÍDA)
        condenser_duty = self.get_condenser_duty("DEST-COL")
        if condenser_duty < 0:
            exergy_condenser = self.calculate_exergy_heat_condenser(condenser_duty)
            total_heat_exergy_output += exergy_condenser
            print(f"    Condensador DEST-COL: {exergy_condenser:.2f} kW (Calor: {condenser_duty:.2f} kW)")

        print(f"\nTOTAL EXERGIA DE CALOR DE ENTRADA: {total_heat_exergy_input:.2f} kW")
        print(f"TOTAL EXERGIA DE CALOR DE SAÍDA: {total_heat_exergy_output:.2f} kW")

        return total_work_exergy, total_heat_exergy_input, total_heat_exergy_output

    def full_exergy_analysis(self):
        """Executa análise exergética completa"""
        print("\n" + "="*60)
        print("ANÁLISE EXERGÉTICA COMPLETA")
        print("="*60)

        try:
            input_streams = ["MKUP-R1", "TGO-1", "MKUP-R3"]

            total_input_exergy = 0.0
            print("\nEXERGIA DE ENTRADA (CORRENTES):")
            for stream in input_streams:
                ex = self.get_stream_exergy(stream)
                print(f"  {stream}: {ex:.2f} kW")
                total_input_exergy += ex
            print(f"TOTAL: {total_input_exergy:.2f} kW")

            self.calculate_pumps_exergy_loss()
            self.calculate_compressors_exergy_loss()
            self.calculate_coolers_exergy_loss()
            self.calculate_mixers_exergy_loss()
            self.calculate_valves_exergy_loss()
            self.calculate_separators_exergy_loss()
            self.calculate_furnaces_exergy_loss()
            self.calculate_heat_exchanger_exergy_loss()
            self.calculate_flash_tanks_exergy_loss()
            self.calculate_columns_exergy_loss()
            self.calculate_reactors_exergy_loss()

            # CÁLCULO CORRIGIDO: Incluir exergias de trabalho e calor separadamente
            total_work_exergy, total_heat_exergy_input, total_heat_exergy_output = self.calculate_total_work_and_heat_exergy()

            # Calcular exergia total de entrada (correntes + trabalho + calor de entrada)
            total_input_exergy_with_work_heat = total_input_exergy + total_work_exergy + total_heat_exergy_input

            output_streams = ["WATER-1", "LIGHTS", "BIO-QAV", "DIESEL", "TAIL-GAS"]
            total_output_exergy = 0.0
            print("\nEXERGIA DE SAÍDA (CORRENTES):")
            for stream in output_streams:
                ex = self.get_stream_exergy(stream)
                print(f"  {stream}: {ex:.2f} kW")
                total_output_exergy += ex

            # CORREÇÃO PRINCIPAL: A exergia do calor removido deve ser somada como valor absoluto às saídas
            total_output_exergy_with_heat = total_output_exergy + total_heat_exergy_output

            # CÁLCULO CORRETO DA PERDA EXERGÉTICA TOTAL DA PLANTA
            # Ex,loss = Ex,in - Ex,out
            # Onde:
            # - Ex,in = correntes de entrada + trabalho + calor fornecido
            # - Ex,out = correntes de saída + calor removido
            total_loss_planta = total_input_exergy_with_work_heat - total_output_exergy_with_heat

            # CÁLCULO CORRETO DA EFICIÊNCIA
            # Eficiência = Ex,out / Ex,in
            if total_input_exergy_with_work_heat > 0:
                efficiency_complete = (total_output_exergy_with_heat / total_input_exergy_with_work_heat) * 100
            else:
                efficiency_complete = 0.0

            # Cálculo da eficiência tradicional (apenas correntes)
            if total_input_exergy > 0:
                efficiency_traditional = (total_output_exergy / total_input_exergy) * 100
            else:
                efficiency_traditional = 0.0

            print("\n" + "="*60)
            print("RESUMO FINAL - BALANÇO EXERGÉTICO COMPLETO")
            print("="*60)
            print("\nEXERGIA DE ENTRADA TOTAL:")
            print(f"  Correntes: {total_input_exergy:.2f} kW")
            print(f"  Trabalho: {total_work_exergy:.2f} kW")
            print(f"  Calor fornecido: {total_heat_exergy_input:.2f} kW")
            print(f"  TOTAL ENTRADA: {total_input_exergy_with_work_heat:.2f} kW")

            print("\nEXERGIA DE SAÍDA TOTAL:")
            print(f"  Correntes: {total_output_exergy:.2f} kW")
            print(f"  Calor removido: {total_heat_exergy_output:.2f} kW")
            print(f"  TOTAL SAÍDA: {total_output_exergy_with_heat:.2f} kW")

            print(f"\nPERDA EXERGÉTICA TOTAL DA PLANTA: {total_loss_planta:.2f} kW")
            print("(Calculada como: Entrada Total - Saída Total)")

            # Verificação do balanço - agora deve ser zero (dentro da precisão numérica)
            balance_difference = total_input_exergy_with_work_heat - total_output_exergy_with_heat - total_loss_planta
            print(f"VERIFICAÇÃO DO BALANÇO: {balance_difference:.2f} kW (deve ser próximo de zero)")

            print("\nEFICIÊNCIAS:")
            print(f"  Eficiência exergética tradicional (apenas correntes): {efficiency_traditional:.2f}%")
            print(f"  Eficiência exergética completa: {efficiency_complete:.2f}%")

            print("\nDETALHAMENTO DAS PERDAS POR EQUIPAMENTO:")
            total_loss_equipamentos = sum(self.results.values())
            for equipment, loss in self.results.items():
                percentual = (loss / total_loss_equipamentos * 100) if total_loss_equipamentos > 0 else 0
                print(f"  {equipment}: {loss:.2f} kW ({percentual:.1f}%)")

            # Adicionar resultados ao dicionário
            self.results['exergia_trabalho_total'] = total_work_exergy
            self.results['exergia_calor_entrada'] = total_heat_exergy_input
            self.results['exergia_calor_saida'] = total_heat_exergy_output
            self.results['exergia_entrada_total'] = total_input_exergy_with_work_heat
            self.results['exergia_saida_total'] = total_output_exergy_with_heat
            self.results['perda_total_planta'] = total_loss_planta
            self.results['eficiencia_tradicional'] = efficiency_traditional
            self.results['eficiencia_completa'] = efficiency_complete
            self.results['balanco_diferenca'] = balance_difference

            return self.results

        except Exception as e:
            print(f"Erro na análise completa: {e}")
            import traceback
            traceback.print_exc()
            return None

def main():
    analyzer = AspenAnalyzer()
    file_path = "C:\\Users\\Usuario\Documents\TCC - Marco Ossips\Simulações\[RTC] - Simulação Análise exergética HVO - Marco.apw"

    if analyzer.connect_to_aspen(file_path):
        try:
            results = analyzer.full_exergy_analysis()
            if results:
                print("\nAnálise concluída com sucesso!")
                print("Resultados finais:", results)
        except Exception as e:
            print(f"Erro durante a análise: {e}")
            import traceback
            traceback.print_exc()
        finally:
            analyzer.close_connection()
    else:
        print("Não foi possível conectar ao Aspen Plus")

if __name__ == '__main__':
    main()