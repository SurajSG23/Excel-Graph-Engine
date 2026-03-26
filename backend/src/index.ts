import { app } from "./app";

const PORT = Number(process.env.PORT || 4100);

app.listen(PORT, () => {
  // Keep startup log concise for local dev and container logs.
  // eslint-disable-next-line no-console
  console.log(`Excel2Graph Pipeline backend running on http://localhost:${PORT}`);
});
