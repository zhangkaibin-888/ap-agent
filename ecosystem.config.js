module.exports = {
  apps: [{
    name: 'invoice-xero-server',
    script: '/root/invoice-xero-bridge/server.py',
    interpreter: 'python3',
    cwd: '/root/invoice-xero-bridge',
    env: {
      PYTHONUNBUFFERED: '1',
    },
    log_file: '/var/log/invoice-xero-server.log',
    error_file: '/var/log/invoice-xero-server-error.log',
    out_file: '/var/log/invoice-xero-server-out.log',
    max_restarts: 10,
    restart_delay: 5000,
  }]
};
