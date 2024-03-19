const Sequelize = require('sequelize');

const connection = new Sequelize('test', 'root', '', {
    dialect: 'mysql',
    host: 'localhost'
});

module.exports = connection;
