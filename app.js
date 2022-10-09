const express = require(`express`);
const methodRouter = require('./methodRoutes');
const cors = require(`cors`);

const app = express();

app.use(cors());

app.use(express.json({ limit: `10kb` }));

app.use(`/api/v1/methods`, methodRouter);

const port = process.env.PORT || 3000;

app.listen(port, () => {
  console.log(`App running on port ${port}`);
});
