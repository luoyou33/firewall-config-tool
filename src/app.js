// 全局变量存储解析的数据
let policyData = null;
let addressData = null;
let serviceData = null;

// 初始化事件监听
document.addEventListener('DOMContentLoaded', () => {
    // 文件上传事件
    document.getElementById('policy-file').addEventListener('change', handleFileUpload('policy'));
    document.getElementById('address-file').addEventListener('change', handleFileUpload('address'));
    document.getElementById('service-file').addEventListener('change', handleFileUpload('service'));
    
    // Tab切换事件
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', switchTab);
    });
    
    document.querySelectorAll('.result-tab').forEach(btn => {
        btn.addEventListener('click', switchResultTab);
    });
    
    // 生成按钮事件
    document.getElementById('generate-btn').addEventListener('click', generateConfig);
    
    // 操作按钮事件
    document.getElementById('copy-btn').addEventListener('click', copyConfig);
    document.getElementById('download-btn').addEventListener('click', downloadConfig);
});

// 文件上传处理
function handleFileUpload(type) {
    return function(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        const statusId = `${type}-status`;
        document.getElementById(statusId).textContent = '解析中...';
        
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);
                
                if (type === 'policy') {
                    policyData = jsonData;
                } else if (type === 'address') {
                    addressData = jsonData;
                } else if (type === 'service') {
                    serviceData = jsonData;
                }
                
                document.getElementById(statusId).textContent = `已读取 ${jsonData.length} 条记录`;
                document.getElementById(statusId).className = 'file-status success';
                
                updateVariableList();
            } catch (error) {
                document.getElementById(statusId).textContent = '解析失败: ' + error.message;
                document.getElementById(statusId).className = 'file-status';
            }
        };
        reader.readAsArrayBuffer(file);
    };
}

// 更新变量列表
function updateVariableList() {
    const variableList = document.getElementById('variable-list');
    let html = '<div class="variable-grid">';
    
    if (policyData && policyData.length > 0) {
        html += '<div class="variable-section"><h4>安全策略变量</h4>';
        Object.keys(policyData[0]).forEach(key => {
            html += `<span class="variable-item">${key}</span>`;
        });
        html += '</div>';
    }
    
    if (addressData && addressData.length > 0) {
        html += '<div class="variable-section"><h4>地址对象变量</h4>';
        Object.keys(addressData[0]).forEach(key => {
            html += `<span class="variable-item">${key}</span>`;
        });
        html += '</div>';
    }
    
    if (serviceData && serviceData.length > 0) {
        html += '<div class="variable-section"><h4>服务对象变量</h4>';
        Object.keys(serviceData[0]).forEach(key => {
            html += `<span class="variable-item">${key}</span>`;
        });
        html += '</div>';
    }
    
    html += '</div>';
    variableList.innerHTML = html;
}

// Tab切换
function switchTab(e) {
    const tabId = e.target.dataset.tab;
    
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    e.target.classList.add('active');
    document.getElementById(`${tabId}-tab`).classList.add('active');
}

// 结果Tab切换
function switchResultTab(e) {
    const resultType = e.target.dataset.result;
    
    document.querySelectorAll('.result-tab').forEach(btn => btn.classList.remove('active'));
    e.target.classList.add('active');
    
    const allConfig = document.getElementById('result-textarea').textContent;
    
    if (resultType === 'all') {
        document.getElementById('result-textarea').textContent = allConfig;
    } else {
        const markers = {
            policy: '安全策略配置',
            address: '地址对象配置',
            service: '服务对象配置'
        };
        const header = `# === ${markers[resultType]} ===`;
        const parts = allConfig.split('# ===');
        const filtered = parts.find(p => p.includes(markers[resultType]));
        document.getElementById('result-textarea').textContent = filtered ? `# ===${filtered}` : '';
    }
}

// 解析IP地址
function parseIPAddress(ipStr) {
    const results = [];
    
    // 处理逗号分隔的多个IP（支持英文逗号和中文逗号）
    if (ipStr.includes(',') || ipStr.includes('，')) {
        const ips = ipStr.split(/[,，]/).map(ip => ip.trim()).filter(ip => ip);
        ips.forEach((ip, index) => {
            // 对每个IP递归处理，支持多种格式
            const subResults = parseSingleIP(ip, index);
            results.push(...subResults);
        });
    }
    // 处理IP范围（使用—或-连接）
    else if (ipStr.includes('—') || ipStr.includes('-')) {
        const regex = /([\d.]+)[—-]([\d.]+)/;
        const match = ipStr.match(regex);
        if (match) {
            results.push(`address 0 range ${match[1]} ${match[2]}`);
        } else {
            results.push(`address 0 ${ipStr} mask 32`);
        }
    }
    // 处理CIDR格式
    else if (ipStr.includes('/')) {
        const parts = ipStr.split('/');
        const ip = parts[0].trim();
        const prefix = parseInt(parts[1]);
        const mask = cidrToMask(prefix);
        results.push(`address 0 ${ip} mask ${mask}`);
    }
    // 单个IP地址
    else {
        results.push(`address 0 ${ipStr} mask 32`);
    }
    
    return results;
}

// 处理单个IP地址
function parseSingleIP(ipStr, baseIndex = 0) {
    const results = [];
    
    // 处理IP范围
    if (ipStr.includes('—') || ipStr.includes('-')) {
        const regex = /([\d.]+)[—-]([\d.]+)/;
        const match = ipStr.match(regex);
        if (match) {
            results.push(`address ${baseIndex} range ${match[1]} ${match[2]}`);
        } else {
            results.push(`address ${baseIndex} ${ipStr} mask 32`);
        }
    }
    // 处理CIDR格式
    else if (ipStr.includes('/')) {
        const parts = ipStr.split('/');
        const ip = parts[0].trim();
        const prefix = parseInt(parts[1]);
        const mask = cidrToMask(prefix);
        results.push(`address ${baseIndex} ${ip} mask ${mask}`);
    }
    // 单个IP地址
    else {
        results.push(`address ${baseIndex} ${ipStr} mask 32`);
    }
    
    return results;
}

// CIDR转子网掩码
function cidrToMask(cidr) {
    const mask = [];
    for (let i = 0; i < 4; i++) {
        const bits = Math.min(8, Math.max(0, cidr - i * 8));
        mask.push(256 - Math.pow(2, 8 - bits));
    }
    return mask.join('.');
}

// 验证IP地址是否在有效范围内
function isValidIP(ip) {
    const parts = ip.split('.');
    if (parts.length !== 4) return false;
    for (let part of parts) {
        const num = parseInt(part);
        if (isNaN(num) || num < 0 || num > 255) return false;
    }
    return true;
}

// 解析单个地址对象
// 返回格式: { type: 'cidr' | 'range' | 'group' | 'single', value: string }
function parseAddressObject(ipStr) {
    // 去除首尾空格
    ipStr = ipStr.trim();
    
    // 检查是否为CIDR格式 (xx.xx.xx.xx/掩码)
    const cidrMatch = ipStr.match(/^([\d.]+)\/(\d{1,2})$/);
    if (cidrMatch) {
        const ip = cidrMatch[1];
        const prefix = parseInt(cidrMatch[2]);
        if (isValidIP(ip) && prefix >= 0 && prefix <= 32) {
            return { type: 'cidr', value: { ip, prefix } };
        }
    }
    
    // 检查是否为地址范围格式 (xx.xx.xx.xx-yy.yy.yy.yy)
    const rangeMatch = ipStr.match(/^([\d.]+)-([\d.]+)$/);
    if (rangeMatch) {
        const startIP = rangeMatch[1];
        const endIP = rangeMatch[2];
        if (isValidIP(startIP) && isValidIP(endIP)) {
            return { type: 'range', value: { startIP, endIP } };
        }
    }
    
    // 检查是否为单个IP地址
    if (isValidIP(ipStr)) {
        return { type: 'single', value: ipStr };
    }
    
    // 否则判断为地址组对象
    return { type: 'group', value: ipStr };
}

// 解析服务端口
function parseServicePort(portStr) {
    const results = [];
    
    // 处理多个服务（逗号分隔，支持英文逗号和中文逗号）
    if (portStr.includes(',') || portStr.includes('，')) {
        const ports = portStr.split(/[,，]/).map(p => p.trim()).filter(p => p);
        ports.forEach((port, index) => {
            const parsed = parseSinglePort(port, index);
            if (parsed) results.push(parsed);
        });
    } else {
        const parsed = parseSinglePort(portStr, 0);
        if (parsed) results.push(parsed);
    }
    
    return results;
}

// 解析单个服务端口
function parseSinglePort(portStr, index) {
    // 提取协议和端口信息
    if (portStr.includes('TCP协议')) {
        const protocol = 'tcp';
        const portMatch = portStr.match(/(\d+)(?:-(.+))?端口/);
        if (portMatch) {
            const startPort = portMatch[1];
            const endPort = portMatch[2] || startPort;
            if (startPort === endPort) {
                return `service ${index} protocol ${protocol} source-port 0 to 65535 destination-port ${startPort}`;
            } else {
                return `service ${index} protocol ${protocol} source-port 0 to 65535 destination-port ${startPort} to ${endPort}`;
            }
        }
    } else if (portStr.includes('UDP协议')) {
        const protocol = 'udp';
        const portMatch = portStr.match(/(\d+)(?:-(.+))?端口/);
        if (portMatch) {
            const startPort = portMatch[1];
            const endPort = portMatch[2] || startPort;
            if (startPort === endPort) {
                return `service ${index} protocol ${protocol} source-port 0 to 65535 destination-port ${startPort}`;
            } else {
                return `service ${index} protocol ${protocol} source-port 0 to 65535 destination-port ${startPort} to ${endPort}`;
            }
        }
    }
    
    return null;
}

// 获取用户自定义的常量配置
function getConstants() {
    return {
        // 安全策略常量 - 前缀
        ruleNamePrefix: document.getElementById('const-rule-name-prefix').value || 'rule name',
        srcZonePrefix: document.getElementById('const-src-zone-prefix').value || 'source-zone',
        dstZonePrefix: document.getElementById('const-dst-zone-prefix').value || 'destination-zone',
        srcAddrPrefix: document.getElementById('const-src-addr-prefix').value || 'source-address',
        dstAddrPrefix: document.getElementById('const-dst-addr-prefix').value || 'destination-address',
        servicePrefix: document.getElementById('const-service-prefix').value || 'service',
        addrSet: document.getElementById('const-addr-set').value || 'address-set',
        action: document.getElementById('const-action').value || 'permit',
        
        // 安全策略常量 - 后缀
        ruleNameSuffix: document.getElementById('const-rule-name-suffix').value || '',
        srcZoneSuffix: document.getElementById('const-src-zone-suffix').value || '',
        dstZoneSuffix: document.getElementById('const-dst-zone-suffix').value || '',
        srcAddrSuffix: document.getElementById('const-src-addr-suffix').value || '',
        dstAddrSuffix: document.getElementById('const-dst-addr-suffix').value || '',
        serviceSuffix: document.getElementById('const-service-suffix').value || '',
        
        // 地址对象常量
        ipSet: document.getElementById('const-ip-set').value || 'ip address-set',
        addrType: document.getElementById('const-addr-type').value || 'type object',
        address: document.getElementById('const-address').value || 'address',
        
        // 服务对象常量
        svcSet: document.getElementById('const-svc-set').value || 'ip service-set',
        serviceItem: document.getElementById('const-service-item').value || 'service',
        
        // 默认配置（可选配）
        enablePolicyLog: document.getElementById('enable-policy-log').checked,
        enableSessionLog: document.getElementById('enable-session-log').checked,
        enableAvProfile: document.getElementById('enable-av-profile').checked,
        enableIpsProfile: document.getElementById('enable-ips-profile').checked,
        policyLog: document.getElementById('const-policy-log').value || 'policy logging',
        sessionLog: document.getElementById('const-session-log').value || 'session logging',
        avProfile: document.getElementById('const-av-profile').value || 'profile av default',
        ipsProfile: document.getElementById('const-ips-profile').value || 'profile ips default'
    };
}

// 生成地址对象配置
function generateAddressConfig() {
    if (!addressData || addressData.length === 0) return '';
    
    const globalConstants = getConstants();
    let config = '# === 地址对象配置 ===\n';
    
    addressData.forEach(item => {
        const ipGroupName = item['IP组名称'] || item['ip_group_name'] || '';
        const ipAddress = item['IP地址'] || item['ip_address'] || '';
        
        // 获取行级别自定义前缀后缀（如果为空则使用全局配置）
        const rowIpSetPrefix = item['地址集前缀'] || globalConstants.ipSet;
        const rowIpSetSuffix = item['地址集后缀'] || '';
        
        if (!ipGroupName || !ipAddress) return;
        
        config += `${rowIpSetPrefix} ${ipGroupName} ${globalConstants.addrType}${rowIpSetSuffix ? ' ' + rowIpSetSuffix : ''}\n`;
        
        const ipLines = parseIPAddress(ipAddress);
        ipLines.forEach(line => {
            config += `  ${line}\n`;
        });
        
        config += '#\n';
    });
    
    return config;
}

// 生成服务对象配置
function generateServiceConfig() {
    if (!serviceData || serviceData.length === 0) return '';
    
    const globalConstants = getConstants();
    let config = '# === 服务对象配置 ===\n';
    
    serviceData.forEach(item => {
        const serviceName = item['服务名称'] || item['service_name'] || '';
        const servicePort = item['服务端口'] || item['service_port'] || '';
        
        // 获取行级别自定义前缀后缀（如果为空则使用全局配置）
        const rowSvcSetPrefix = item['服务集前缀'] || globalConstants.svcSet;
        const rowSvcSetSuffix = item['服务集后缀'] || '';
        
        if (!serviceName || !servicePort) return;
        
        config += `${rowSvcSetPrefix} ${serviceName} ${globalConstants.addrType}${rowSvcSetSuffix ? ' ' + rowSvcSetSuffix : ''}\n`;
        
        const serviceLines = parseServicePort(servicePort);
        serviceLines.forEach(line => {
            config += `  ${line}\n`;
        });
        
        config += '#\n';
    });
    
    return config;
}

// 生成安全策略配置
function generatePolicyConfig() {
    if (!policyData || policyData.length === 0) return '';
    
    const globalConstants = getConstants();
    let config = '# === 安全策略配置 ===\n';
    
    policyData.forEach(item => {
        const policyName = item['安全策略名称'] || item['policy_name'] || '';
        const srcZone = item['源安全区域'] || item['src_zone'] || '';
        const srcIP = item['源IP地址'] || item['src_ip'] || '';
        const dstZone = item['目的安全区域'] || item['dst_zone'] || '';
        const dstIP = item['目的IP地址'] || item['dst_ip'] || '';
        const service = item['服务'] || item['service'] || '';
        const action = item['动作'] || item['action'] || '允许';
        
        // 获取行级别自定义前缀后缀（如果为空则使用全局配置）
        const rowConstants = {
            ruleNamePrefix: item['规则名称前缀'] || globalConstants.ruleNamePrefix,
            ruleNameSuffix: item['规则名称后缀'] || globalConstants.ruleNameSuffix,
            srcZonePrefix: item['源区域前缀'] || globalConstants.srcZonePrefix,
            srcZoneSuffix: item['源区域后缀'] || globalConstants.srcZoneSuffix,
            dstZonePrefix: item['目的区域前缀'] || globalConstants.dstZonePrefix,
            dstZoneSuffix: item['目的区域后缀'] || globalConstants.dstZoneSuffix,
            srcAddrPrefix: item['源地址前缀'] || globalConstants.srcAddrPrefix,
            srcAddrSuffix: item['源地址后缀'] || globalConstants.srcAddrSuffix,
            dstAddrPrefix: item['目的地址前缀'] || globalConstants.dstAddrPrefix,
            dstAddrSuffix: item['目的地址后缀'] || globalConstants.dstAddrSuffix,
            servicePrefix: item['服务前缀'] || globalConstants.servicePrefix,
            serviceSuffix: item['服务后缀'] || globalConstants.serviceSuffix,
            addrSet: globalConstants.addrSet,
            action: globalConstants.action,
            enablePolicyLog: globalConstants.enablePolicyLog,
            enableSessionLog: globalConstants.enableSessionLog,
            enableAvProfile: globalConstants.enableAvProfile,
            enableIpsProfile: globalConstants.enableIpsProfile,
            policyLog: globalConstants.policyLog,
            sessionLog: globalConstants.sessionLog,
            avProfile: globalConstants.avProfile,
            ipsProfile: globalConstants.ipsProfile
        };
        
        if (!policyName) return;
        
        // 规则名称：前缀 + 变量 + 后缀（使用行级别配置）
        config += `${rowConstants.ruleNamePrefix} ${policyName}${rowConstants.ruleNameSuffix ? ' ' + rowConstants.ruleNameSuffix : ''}\n`;
        
        // 处理源区域（可能有多个，支持中文逗号和英文逗号）
        const srcZones = srcZone.split(/[,，]/).map(z => z.trim()).filter(z => z && z !== 'any');
        if (srcZones.length > 0) {
            srcZones.forEach(zone => {
                config += `  ${rowConstants.srcZonePrefix} ${zone}${rowConstants.srcZoneSuffix ? ' ' + rowConstants.srcZoneSuffix : ''}\n`;
            });
        } else if (srcZone === 'any') {
            config += `  ${rowConstants.srcZonePrefix} any${rowConstants.srcZoneSuffix ? ' ' + rowConstants.srcZoneSuffix : ''}\n`;
        }
        
        // 处理目的区域（支持中文逗号和英文逗号）
        const dstZones = dstZone.split(/[,，]/).map(z => z.trim()).filter(z => z && z !== 'any');
        if (dstZones.length > 0) {
            dstZones.forEach(zone => {
                config += `  ${rowConstants.dstZonePrefix} ${zone}${rowConstants.dstZoneSuffix ? ' ' + rowConstants.dstZoneSuffix : ''}\n`;
            });
        } else if (dstZone === 'any') {
            config += `  ${rowConstants.dstZonePrefix} any${rowConstants.dstZoneSuffix ? ' ' + rowConstants.dstZoneSuffix : ''}\n`;
        }
        
        // 处理源地址（支持中文逗号和英文逗号）
        const srcIPs = srcIP.split(/[,，]/).map(ip => ip.trim()).filter(ip => ip && ip !== 'any');
        if (srcIPs.length > 0) {
            srcIPs.forEach(ip => {
                const parsed = parseAddressObject(ip);
                let line = '';
                
                switch (parsed.type) {
                    case 'cidr':
                        // CIDR格式：前缀 xx.xx.xx.xx mask 掩码
                        const cidrMask = cidrToMask(parsed.value.prefix);
                        line = `${rowConstants.srcAddrPrefix} ${parsed.value.ip} mask ${cidrMask}`;
                        break;
                    case 'range':
                        // 地址范围格式：前缀 range xx.xx.xx.xx yy.yy.yy.yy
                        line = `${rowConstants.srcAddrPrefix} range ${parsed.value.startIP} ${parsed.value.endIP}`;
                        break;
                    case 'single':
                        // 单个IP：前缀 xx.xx.xx.xx mask 255.255.255.255
                        line = `${rowConstants.srcAddrPrefix} ${parsed.value} mask 255.255.255.255`;
                        break;
                    case 'group':
                    default:
                        // 地址组对象：前缀 address-set 地址组名称
                        line = `${rowConstants.srcAddrPrefix} ${rowConstants.addrSet} ${parsed.value}`;
                        break;
                }
                
                config += `  ${line}${rowConstants.srcAddrSuffix ? ' ' + rowConstants.srcAddrSuffix : ''}\n`;
            });
        } else if (srcIP === 'any') {
            config += `  ${rowConstants.srcAddrPrefix} any${rowConstants.srcAddrSuffix ? ' ' + rowConstants.srcAddrSuffix : ''}\n`;
        }
        
        // 处理目的地址（支持中文逗号和英文逗号）
        const dstIPs = dstIP.split(/[,，]/).map(ip => ip.trim()).filter(ip => ip && ip !== 'any');
        if (dstIPs.length > 0) {
            dstIPs.forEach(ip => {
                const parsed = parseAddressObject(ip);
                let line = '';
                
                switch (parsed.type) {
                    case 'cidr':
                        // CIDR格式：前缀 xx.xx.xx.xx mask 掩码
                        const cidrMask = cidrToMask(parsed.value.prefix);
                        line = `${rowConstants.dstAddrPrefix} ${parsed.value.ip} mask ${cidrMask}`;
                        break;
                    case 'range':
                        // 地址范围格式：前缀 range xx.xx.xx.xx yy.yy.yy.yy
                        line = `${rowConstants.dstAddrPrefix} range ${parsed.value.startIP} ${parsed.value.endIP}`;
                        break;
                    case 'single':
                        // 单个IP：前缀 xx.xx.xx.xx mask 255.255.255.255
                        line = `${rowConstants.dstAddrPrefix} ${parsed.value} mask 255.255.255.255`;
                        break;
                    case 'group':
                    default:
                        // 地址组对象：前缀 address-set 地址组名称
                        line = `${rowConstants.dstAddrPrefix} ${rowConstants.addrSet} ${parsed.value}`;
                        break;
                }
                
                config += `  ${line}${rowConstants.dstAddrSuffix ? ' ' + rowConstants.dstAddrSuffix : ''}\n`;
            });
        } else if (dstIP === 'any') {
            config += `  ${rowConstants.dstAddrPrefix} any${rowConstants.dstAddrSuffix ? ' ' + rowConstants.dstAddrSuffix : ''}\n`;
        }
        
        // 处理服务
        if (service && service !== 'any') {
            config += `  ${rowConstants.servicePrefix} ${service}${rowConstants.serviceSuffix ? ' ' + rowConstants.serviceSuffix : ''}\n`;
        }
        
        // 默认配置（使用全局配置）
        if (rowConstants.enablePolicyLog && rowConstants.policyLog) {
            config += `  ${rowConstants.policyLog}\n`;
        }
        if (rowConstants.enableSessionLog && rowConstants.sessionLog) {
            config += `  ${rowConstants.sessionLog}\n`;
        }
        if (rowConstants.enableAvProfile && rowConstants.avProfile) {
            config += `  ${rowConstants.avProfile}\n`;
        }
        if (rowConstants.enableIpsProfile && rowConstants.ipsProfile) {
            config += `  ${rowConstants.ipsProfile}\n`;
        }
        
        // 动作（优先使用Excel中的动作，否则使用常量配置）
        let actionCmd = action === '允许' ? rowConstants.action : 
                       (action === '拒绝' ? 'deny' : rowConstants.action);
        config += `  action ${actionCmd}\n`;
        
        config += '#\n';
    });
    
    return config;
}

// 生成配置
function generateConfig() {
    let config = '';
    
    // 生成地址对象配置
    if (document.getElementById('gen-address').checked) {
        config += generateAddressConfig();
    }
    
    // 生成服务对象配置
    if (document.getElementById('gen-service').checked) {
        if (config) config += '\n';
        config += generateServiceConfig();
    }
    
    // 生成安全策略配置
    if (document.getElementById('gen-policy').checked) {
        if (config) config += '\n';
        config += generatePolicyConfig();
    }
    
    if (config === '') {
        alert('请至少上传一个规划表并选择生成选项');
        return;
    }
    
    document.getElementById('result-textarea').textContent = config;
    
    // 切换到全部配置标签
    document.querySelectorAll('.result-tab').forEach(btn => btn.classList.remove('active'));
    document.querySelector('.result-tab[data-result="all"]').classList.add('active');
}

// 复制配置
function copyConfig() {
    const textarea = document.getElementById('result-textarea');
    textarea.select();
    document.execCommand('copy');
    
    const btn = document.getElementById('copy-btn');
    const originalText = btn.textContent;
    btn.textContent = '已复制!';
    btn.style.background = '#28a745';
    
    setTimeout(() => {
        btn.textContent = originalText;
        btn.style.background = '#6c757d';
    }, 2000);
}

// 下载配置
function downloadConfig() {
    const content = document.getElementById('result-textarea').textContent;
    if (!content) {
        alert('请先生成配置');
        return;
    }
    
    const blob = new Blob([content], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `firewall_config_${new Date().toISOString().slice(0, 10)}.txt`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// 添加拖拽支持
['policy', 'address', 'service'].forEach(type => {
    const label = document.querySelector(`#${type}-group .file-label`);
    
    label.addEventListener('dragover', (e) => {
        e.preventDefault();
        label.classList.add('dragover');
    });
    
    label.addEventListener('dragleave', () => {
        label.classList.remove('dragover');
    });
    
    label.addEventListener('drop', (e) => {
        e.preventDefault();
        label.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            document.getElementById(`${type}-file`).files = files;
            document.getElementById(`${type}-file`).dispatchEvent(new Event('change'));
        }
    });
});

// 下载模板函数
function downloadTemplate(type) {
    let headers = [];
    let fileName = '';
    let examples = [];
    
    switch(type) {
        case 'policy':
            headers = [
                '安全策略名称', '源安全区域', '源IP地址', '目的安全区域', '目的IP地址', '服务', '动作',
                '规则名称前缀', '规则名称后缀',
                '源区域前缀', '源区域后缀',
                '源地址前缀', '源地址后缀',
                '目的区域前缀', '目的区域后缀',
                '目的地址前缀', '目的地址后缀',
                '服务前缀', '服务后缀'
            ];
            fileName = '安全策略规划表模板.xlsx';
            examples = [
                ['安全策略01', '办公zone', 'IP组01', '服务器zone', 'IP组02', 'TCP443', '允许', '', '', '', '', '', '', '', '', '', '', ''],
                ['安全策略02', '办公zone', 'IP组03', '互联网zone', 'any', 'any', '允许', '', '', '', '', '', '', '', '', '', '', ''],
                ['安全策略03', 'any', 'IP组04', 'any', '192.168.3.10', 'TCP1000-1005', '拒绝', '', '', '', '', '', '', '', '', '', '', '']
            ];
            break;
        case 'address':
            headers = ['IP组名称', 'IP地址', '地址集前缀', '地址集后缀'];
            fileName = '地址对象规划表模板.xlsx';
            examples = [
                ['IP组01', '192.168.1.1,192.168.1.2', '', ''],
                ['IP组02', '172.16.1.0/24', '', ''],
                ['IP组03', '10.0.0.1-10.0.0.10', '', '']
            ];
            break;
        case 'service':
            headers = ['服务名称', '服务端口', '服务集前缀', '服务集后缀'];
            fileName = '服务对象规划表模板.xlsx';
            examples = [
                ['TCP443', 'TCP协议443端口', '', ''],
                ['UDP123', 'UDP协议123端口', '', ''],
                ['TCP1000-1005', 'TCP协议1000-1005端口', '', '']
            ];
            break;
        default:
            return;
    }
    
    // 创建工作簿和工作表
    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...examples]);
    
    // 设置列宽
    worksheet['!cols'] = headers.map(() => ({ wch: 15 }));
    
    // 创建工作簿
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '模板');
    
    // 生成Excel文件并下载
    XLSX.writeFile(workbook, fileName);
}
